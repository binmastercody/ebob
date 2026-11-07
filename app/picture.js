/* Picture guide: Copy PNG, Download PNG and Email guide.
   A hidden iframe (?picture=1) steps through the lesson and shoots each screen
   with the red circle on it. The parent composes a 1200 px walkthrough.
   Works from file:// and from GitHub Pages. No server or CDN. */
(function () {
    "use strict";
    var LS = window.EbobLessons;
    if (!LS) return;
    var SITE = "https://binmastercody.github.io/ebob/";
    var W = 1200, MAX_H = 13500, PART_MAX = 10, MAIL_MAX = 1800;
    var HEAD = 124, RED_H = 5, HEAD_H = HEAD + RED_H;
    var FONT = '"DM Sans","Outfit","Segoe UI",Arial,sans-serif';
    var C = { header: "#282C38", red: "#C0202A", violet: "#5B3FA6", bg: "#F3F4F6", text: "#282C38", body: "#363B47", white: "#FFFFFF", target: "#E0201C" };

    function qs() { return new URLSearchParams(location.search); }
    function inFrame() { return qs().get("picture") === "1"; }
    function hrefOf(l) {
        return l.page === "installer" ? "installer.html?lesson=" + l.id
            : l.kind === "doc" ? "doc.html?lesson=" + l.id
            : l.scenario ? "emulator.html?tutorial=" + l.scenario + "&fresh=1"
            : "emulator.html?tutorial=free-mode&lesson=" + l.id;
    }
    function plain(html) {
        var d = document.createElement("div");
        d.innerHTML = String(html == null ? "" : html).replace(/<li[^>]*>/gi, "\n- ").replace(/<\/(p|h\d|ol|ul|dl|div)>/gi, "\n").replace(/<br\s*\/?>/gi, "\n");
        return (d.textContent || "").replace(/[ \t]+/g, " ").replace(/\n\s*\n+/g, "\n").replace(/[\u2014\u2013]/g, ", ").trim();
    }
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
    function blobToData(b) {
        return new Promise(function (res, rej) {
            var r = new FileReader();
            r.onload = function () { res(r.result); };
            r.onerror = function () { rej(new Error("could not read picture")); };
            r.readAsDataURL(b);
        });
    }
    function loadImg(src) {
        return new Promise(function (res) {
            if (!src) return res(null);
            var im = new Image();
            im.onload = function () { res(im); };
            im.onerror = function () { res(null); };
            im.src = src;
        });
    }
    function fileName(id, i, n) {
        return "binventory-guide-" + id + (n > 1 ? "-part-" + (i + 1) + "-of-" + n : "") + ".png";
    }

    function rewriteAssets(doc) {
        var assets = window.EbobPictureAssets || {};
        var keys = Object.keys(assets);
        function match(src) {
            if (!src) return null;
            for (var i = 0; i < keys.length; i++) {
                if (src.indexOf(keys[i]) !== -1) return assets[keys[i]];
            }
            return null;
        }
        var imgs = doc.querySelectorAll("img");
        for (var i = 0; i < imgs.length; i++) {
            var d = match(imgs[i].getAttribute("src") || imgs[i].src);
            if (d) imgs[i].src = d;
        }
    }
    async function shoot(win) {
        var doc = win.document;
        rewriteAssets(doc);
        doc.documentElement.classList.add("eg-picture-mode");
        var api = win.modernScreenshot || window.modernScreenshot;
        if (api && api.domToBlob) {
            return api.domToBlob(doc.documentElement, {
                width: win.innerWidth || 1200,
                height: win.innerHeight || 800,
                scale: 1,
                backgroundColor: "#E4E7EC",
                filter: function (el) {
                    if (!el || !el.classList) return true;
                    if (el.classList.contains("eg-pic-dlg") || el.id === "egPicMenu") return false;
                    if (el.getAttribute && el.getAttribute("data-share")) return false;
                    return true;
                }
            });
        }
        throw new Error("screenshot library missing");
    }

    function stepInfoFromLesson(l, i, win) {
        if (l.kind === "doc" && l.doc && l.doc[i]) {
            return { title: l.doc[i].title, text: plain(l.doc[i].html), ring: false };
        }
        if (l.steps && l.steps[i]) {
            var st = l.steps[i];
            var msg = typeof st.msg === "function" ? st.msg() : st.msg;
            return { title: st.title, text: plain(msg), ring: st.ring !== false };
        }
        if (win && win.__ebobNativeTutorial) {
            var ns = win.__ebobNativeTutorial.state();
            var step = ns && ns.steps && ns.steps[i];
            if (step) return { title: step.title, text: plain(step.message), ring: true };
        }
        return { title: "Step " + (i + 1), text: "", ring: false };
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
        var opts = { bubbles: true, cancelable: true, view: el.ownerDocument.defaultView, button: right ? 2 : 0, buttons: right ? 2 : 1 };
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
            var m = (step.message || "").match(/["\u201c]([^"\u201d]+)["\u201d]/);
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

    async function driveSelf(progress) {
        progress = progress || function () {};
        var sp = qs();
        var id = sp.get("lesson");
        var tutorial = sp.get("tutorial");
        var l = id ? LS.byId(id) : null;
        if (!l && tutorial && tutorial !== "free-mode") {
            LS.lessons.forEach(function (x) { if (x.scenario === tutorial) l = x; });
        }
        if (!l) throw new Error("no lesson in this page");
        if (window.EbobDisplay && window.EbobDisplay.PACES && window.EbobDisplay.PACES.picture) {
            try { window.EbobDisplay.set("pace", "picture"); window.EbobDisplay.set("size", 1); window.EbobDisplay.set("contrast", false); } catch (e) {}
        }
        document.documentElement.classList.add("eg-picture-mode");
        var n = l.stepCount;
        var out = [];
        if (l.kind === "doc") {
            await until(function () { return document.getElementById("docNext") && document.getElementById("docStepTitle"); });
            for (var i = 0; i < n; i++) {
                progress("Building picture guide: step " + (i + 1) + " of " + n);
                await sleep(250);
                var info = stepInfoFromLesson(l, i, window);
                var blob = await shoot(window);
                out.push({ title: info.title, text: info.text, ring: false, ringBox: null, dataUrl: await blobToData(blob) });
                if (i < n - 1) document.getElementById("docNext").click();
            }
            return out;
        }
        if (l.scenario) {
            await until(function () { return window.__ebobNativeTutorial && window.__ebobNativeTutorial.state() && window.__ebobNativeTutorial.state().total; }, 90000);
            for (var j = 0; j < n; j++) {
                progress("Building picture guide: step " + (j + 1) + " of " + n);
                await until(function () {
                    var s = window.__ebobNativeTutorial.state();
                    return s && s.index === j;
                }, 30000);
                await sleep(400);
                var ninfo = stepInfoFromLesson(l, j, window);
                var circle = document.getElementById("ebobTutorialCircle");
                var cr = circle && circle.getBoundingClientRect();
                var ring = !!(cr && cr.width > 2);
                var ringBox = ring ? { l: cr.left, t: cr.top, r: cr.right, b: cr.bottom } : null;
                var nblob = await shoot(window);
                out.push({ title: ninfo.title, text: ninfo.text, ring: ring, ringBox: ringBox, dataUrl: await blobToData(nblob) });
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
                    await until(function () {
                        var s = window.__ebobNativeTutorial.state();
                        return s && (s.index > j || s.completed);
                    }, 20000);
                }
            }
            if (out.length !== n) throw new Error("step count mismatch: got " + out.length + " expected " + n);
            return out;
        }
        var G = window.EbobGuide;
        await until(function () { var s = G && G.state && G.state(); return s && s.phase && s.phase !== "loading" && s.total; });
        for (var k = 0; k < n; k++) {
            progress("Building picture guide: step " + (k + 1) + " of " + n);
            await until(function () {
                var s = G.state();
                return s && s.idx === k && (s.phase === "ready" || s.phase === "done" || s.finished);
            });
            window.__ebobPictureLock = true;
            await sleep(400);
            var infoG = stepInfoFromLesson(l, k, window);
            var layout = G.layout && G.layout();
            var ringG = !!(layout && layout.ring);
            var ringBox = layout && layout.ring ? { l: layout.ring.l, t: layout.ring.t, r: layout.ring.r, b: layout.ring.b } : null;
            if (l.steps && l.steps[k] && l.steps[k].ring === false) { ringG = false; ringBox = null; }
            var gblob = await shoot(window);
            out.push({ title: infoG.title, text: infoG.text, ring: ringG, ringBox: ringBox, dataUrl: await blobToData(gblob) });
            window.__ebobPictureLock = false;
            if (k < n - 1) {
                if (G.act) G.act();
                await until(function () {
                    var s = G.state();
                    return s && (s.idx > k || s.finished);
                });
            }
        }
        if (out.length !== n) throw new Error("step count mismatch: got " + out.length + " expected " + n);
        return out;
    }

    function wrap(ctx, text, maxW) {
        var out = [];
        String(text || "").split("\n").forEach(function (para) {
            var words = para.split(" "), line = "";
            words.forEach(function (w) {
                var t = line ? line + " " + w : w;
                if (ctx.measureText(t).width > maxW && line) { out.push(line); line = w; } else line = t;
            });
            if (line || !words.length) out.push(line);
        });
        return out.filter(function (x, i, a) { return x || i === 0 || a.length === 1; });
    }

    async function compose(l, shots) {
        var logoSrc = (window.EbobPictureAssets || {})["app/binmaster-logo-white.png"];
        var logo = await loadImg(logoSrc);
        var ctxm = document.createElement("canvas").getContext("2d");
        var items = [];
        for (var i = 0; i < shots.length; i++) {
            var im = await loadImg(shots[i].dataUrl);
            ctxm.font = "700 24px " + FONT;
            var tl = wrap(ctxm, shots[i].title || "", W - 130);
            ctxm.font = "17px " + FONT;
            var ml = shots[i].text ? wrap(ctxm, shots[i].text, W - 130) : [];
            var sh = im ? Math.round(im.height * (W / im.width)) : 0;
            var h = 28 + tl.length * 30 + (ml.length ? 8 + ml.length * 24 : 0) + 16 + sh + 28;
            items.push({ s: shots[i], im: im, tl: tl, ml: ml, sh: sh, h: h });
        }
        var groups = [], cur = [], ch = HEAD_H + 48;
        items.forEach(function (it) {
            if (cur.length && (cur.length >= PART_MAX || ch + it.h > MAX_H)) {
                groups.push(cur); cur = []; ch = HEAD_H + 48;
            }
            cur.push(it); ch += it.h;
        });
        if (cur.length) groups.push(cur);

        function draw(group, part, parts) {
            var body = 0; group.forEach(function (it) { body += it.h; });
            var H = HEAD_H + 36 + body + 36;
            var cv = document.createElement("canvas");
            cv.width = W; cv.height = H;
            var x = cv.getContext("2d");
            x.fillStyle = C.bg; x.fillRect(0, 0, W, H);
            x.fillStyle = C.header; x.fillRect(0, 0, W, HEAD);
            x.fillStyle = C.red; x.fillRect(0, HEAD, W, RED_H);
            if (logo) {
                var lh = 28, lw = Math.round(logo.width * (lh / logo.height));
                x.drawImage(logo, 28, 30, lw, lh);
            } else {
                x.fillStyle = C.white; x.font = "700 18px " + FONT; x.fillText("BinMaster", 28, 50);
            }
            x.fillStyle = C.white; x.font = "600 16px " + FONT; x.textAlign = "right";
            x.fillText("Binventory Workstation 1.0", W - 28, 50);
            x.textAlign = "left";
            x.font = "700 32px " + FONT;
            var tt = wrap(x, l.title, W - 56)[0] || l.title;
            x.fillText(tt, 28, 100);
            var y = HEAD_H + 36;
            group.forEach(function (it, gi) {
                var num = shots.indexOf(it.s) + 1;
                x.fillStyle = C.violet;
                x.beginPath(); x.arc(64, y + 24, 24, 0, Math.PI * 2); x.fill();
                x.fillStyle = C.white; x.font = "700 16px " + FONT; x.textAlign = "center";
                x.fillText(String(num), 64, y + 30); x.textAlign = "left";
                x.fillStyle = C.text; x.font = "700 24px " + FONT;
                it.tl.forEach(function (line, j) { x.fillText(line, 100, y + 24 + j * 30); });
                y += 28 + it.tl.length * 30;
                if (it.ml.length) {
                    y += 8; x.fillStyle = C.body; x.font = "17px " + FONT;
                    it.ml.forEach(function (line, j) { x.fillText(line, 100, y + 16 + j * 24); });
                    y += it.ml.length * 24;
                }
                y += 16;
                if (it.im) {
                    x.drawImage(it.im, 0, y, W, it.sh);
                    var box = it.s.ringBox;
                    if (box) {
                        var sx = W / it.im.width, sy = it.sh / it.im.height;
                        var cx = ((box.l + box.r) / 2) * sx;
                        var cy = y + ((box.t + box.b) / 2) * sy;
                        var rx = Math.max(22, (box.r - box.l) / 2 * sx + 14);
                        var ry = Math.max(22, (box.b - box.t) / 2 * sy + 14);
                        x.save();
                        x.strokeStyle = "#E0201C";
                        x.lineWidth = 6;
                        x.beginPath();
                        x.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
                        x.stroke();
                        x.restore();
                    }
                }
                y += it.sh + 28;
            });
            return cv;
        }
        function toBlob(cv) {
            return new Promise(function (res) { cv.toBlob(function (b) { res(b); }, "image/png"); });
        }
        var parts = [];
        for (var p = 0; p < groups.length; p++) {
            parts.push({ blob: await toBlob(draw(groups[p], p + 1, groups.length)), canvas: null });
        }
        return parts;
    }

    async function captureLesson(l, progress) {
        if (inFrame()) return driveSelf(progress);
        return new Promise(function (resolve, reject) {
            var iframe = document.createElement("iframe");
            iframe.setAttribute("title", "Building picture guide");
            iframe.setAttribute("aria-hidden", "true");
            iframe.style.cssText = "position:fixed;left:0;top:0;width:1440px;height:900px;border:0;opacity:0;pointer-events:none;z-index:-1;";
            var url = hrefOf(l);
            url += (url.indexOf("?") >= 0 ? "&" : "?") + "picture=1";
            var done = false;
            function finish(err, result) {
                if (done) return; done = true;
                window.removeEventListener("message", onMsg);
                setTimeout(function () { if (iframe.parentNode) iframe.remove(); }, 80);
                if (err) reject(err); else resolve(result);
            }
            function onMsg(e) {
                if (e.source !== iframe.contentWindow) return;
                var d = e.data || {};
                if (d.type === "ebob-pic-progress" && d.text) progress(d.text);
                if (d.type === "ebob-pic-done") finish(d.err ? new Error(d.err) : null, d.steps);
            }
            window.addEventListener("message", onMsg);
            iframe.onload = function () {
                try {
                    var w = iframe.contentWindow;
                    if (w && w.EbobPicture && typeof w.EbobPicture.driveSelf === "function") {
                        w.EbobPicture.driveSelf(progress).then(function (steps) { finish(null, steps); }, function (e) { finish(e); });
                        return;
                    }
                } catch (err) {}
                try { iframe.contentWindow.postMessage({ type: "ebob-pic-run" }, "*"); } catch (e2) { finish(e2); }
            };
            iframe.src = url;
            document.body.appendChild(iframe);
            setTimeout(function () { finish(new Error("picture guide timed out")); }, 170000);
        });
    }

    async function build(id, progress) {
        var l = LS.byId(id);
        if (!l) throw new Error("unknown lesson " + id);
        progress = progress || function () {};
        var shots = await captureLesson(l, progress);
        if (!shots || shots.length !== l.stepCount) {
            throw new Error("step count mismatch: got " + (shots && shots.length) + " expected " + l.stepCount);
        }
        var parts = await compose(l, shots);
        return {
            lesson: l,
            steps: shots.map(function (s) { return { title: s.title, text: s.text }; }),
            rings: shots.map(function (s) { return !!s.ring; }),
            parts: parts
        };
    }

    function mailtoUrl(l, steps, copied) {
        var subject = "Binventory guide: " + l.title;
        var link = "Online lesson: " + SITE + hrefOf(l);
        var note = copied
            ? "The picture guide is on your clipboard. Click in this email and press Ctrl+V to paste it. It is also in your Downloads folder."
            : "The picture guide is in your Downloads folder.";
        function pack(lines, extra) {
            return ["Hello,", ""].concat(lines).concat(["", link, "", note]).concat(extra || []).join("\r\n");
        }
        function urlFor(lines, extra) {
            return "mailto:?subject=" + encodeURIComponent(subject) + "&body=" + encodeURIComponent(pack(lines, extra));
        }
        var listed = [];
        for (var i = 0; i < steps.length; i++) {
            var s = steps[i];
            var title = typeof s === "string" ? s : (s && s.title) || ("Step " + (i + 1));
            var line = (i + 1) + ". " + String(title).replace(/[\r\n]+/g, " ");
            if (line.length > 110) line = line.slice(0, 107).replace(/\s+\S*$/, "") + "...";
            var extra = (i + 1 < steps.length) ? ["See the picture for all steps."] : [];
            if (urlFor(listed.concat([line]), extra).length > MAIL_MAX) break;
            listed.push(line);
        }
        var extra = listed.length < steps.length ? ["See the picture for all steps."] : [];
        var url = urlFor(listed, extra);
        while (url.length > MAIL_MAX && listed.length) {
            listed.pop();
            extra = ["See the picture for all steps."];
            url = urlFor(listed, extra);
        }
        return url;
    }

    function downloadOne(blob, name) {
        var a = document.createElement("a");
        a.className = "eg-ui";
        a.href = URL.createObjectURL(blob);
        a.download = name;
        a.style.cssText = "position:fixed;left:-9999px;top:0;width:1px;height:1px;";
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    }
    function downloadAll(id, parts) {
        parts.forEach(function (p, i) { downloadOne(p.blob, fileName(id, i, parts.length)); });
    }
    function openMail(url) {
        var a = document.createElement("a");
        a.className = "eg-ui";
        a.href = url;
        a.style.cssText = "position:fixed;left:-9999px;top:0;";
        document.body.appendChild(a);
        a.click();
        a.remove();
    }

    var last = null, opener = null, dlg = null;

    function setStatus(text) {
        var s = document.querySelector(".eg-pic-status");
        if (s) s.textContent = text;
    }
    function closeDlg() {
        if (dlg && dlg.parentNode) dlg.remove();
        dlg = null;
        if (opener) { try { opener.focus(); } catch (e) {} }
    }
    function showDialog(text) {
        if (!dlg) {
            dlg = document.createElement("div");
            dlg.className = "eg-ui eg-pic-dlg";
            dlg.innerHTML = '<div class="eg-pic-card" role="dialog" aria-modal="true" aria-labelledby="egPicDlgTitle">' +
                '<h2 id="egPicDlgTitle" class="eg-pic-dlg-title">Picture guide</h2>' +
                '<p class="eg-pic-status" role="status"></p>' +
                '<div class="eg-pic-actions"></div></div>';
            document.body.appendChild(dlg);
        }
        setStatus(text);
        return dlg;
    }
    function finishDialog(state, id, parts) {
        last = state;
        window.EbobPicture.last = last;
        showDialog(state.message);
        var actions = dlg.querySelector(".eg-pic-actions");
        actions.innerHTML = "";
        var showSave = !(state.copied && parts.length === 1 && !state.forceSave);
        var save = null;
        if (showSave) {
            save = document.createElement("button");
            save.type = "button";
            save.className = "eg-btn eg-pic-save";
            save.textContent = parts.length > 1 ? "Save pictures" : "Save picture";
            save.addEventListener("click", function () {
                downloadAll(id, parts);
            });
            actions.appendChild(save);
        }
        var ok = document.createElement("button");
        ok.type = "button";
        ok.className = "eg-btn eg-btn--ghost eg-pic-cancel";
        ok.textContent = "OK";
        ok.addEventListener("click", closeDlg);
        actions.appendChild(ok);
        dlg.querySelector(".eg-pic-card").addEventListener("keydown", function (e) {
            if (e.key !== "Tab") return;
            var list = [save, ok].filter(Boolean);
            if (!list.length) return;
            var i = list.indexOf(document.activeElement);
            e.preventDefault();
            var next = e.shiftKey ? (i <= 0 ? list.length - 1 : i - 1) : (i >= list.length - 1 ? 0 : i + 1);
            list[next].focus();
        });
        setTimeout(function () { (save || ok).focus(); }, 0);
    }

    async function copyFirst(blob) {
        if (!navigator.clipboard || !navigator.clipboard.write || !window.ClipboardItem) {
            var err = new Error("This browser cannot put pictures on the clipboard.");
            err.code = "noclip";
            throw err;
        }
        try {
            await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        } catch (e) {
            var err2 = new Error("The browser did not allow copying.");
            err2.code = "denied";
            throw err2;
        }
    }

    async function run(act, id) {
        closeMenu();
        var l = LS.byId(id);
        if (!l) return;
        showDialog("Building picture guide...");
        var buildP = build(id, setStatus);
        /* start the clipboard write during the click so the browser keeps permission while the picture builds */
        var copyP = null;
        if (act !== "download") {
            if (navigator.clipboard && navigator.clipboard.write && window.ClipboardItem) {
                try {
                    copyP = navigator.clipboard.write([
                        new ClipboardItem({
                            "image/png": buildP.then(function (m) { return m.parts[0].blob; })
                        })
                    ]).then(function () { return true; }, function () { return false; });
                } catch (e) {
                    copyP = Promise.resolve(false);
                }
            }
        }
        try {
            var m = await buildP;
            var parts = m.parts;
            var copied, copyError;
            if (act === "download") {
                finishDialog({
                    message: parts.length > 1
                        ? "Done. Click Save pictures to put them in your Downloads folder."
                        : "Done. Click Save picture to put it in your Downloads folder.",
                    parts: parts.length,
                    forceSave: true
                }, id, parts);
                return;
            }
            if (copyP) {
                copied = await copyP;
                if (!copied) copyError = "The browser did not allow copying.";
            } else {
                copied = false;
                copyError = "This browser cannot put pictures on the clipboard.";
            }
            if (act === "email") {
                window.EbobPicture.openMail(mailtoUrl(l, m.steps, !!copied));
                var em = { message: "Done. Your email program should open with a new message.", parts: parts.length, copied: copied, forceSave: true };
                if (copyError) em.copyError = copyError;
                finishDialog(em, id, parts);
                return;
            }
            if (copied) {
                finishDialog({
                    message: parts.length > 1
                        ? "Done. Part 1 of " + parts.length + " is on your clipboard. Paste it with Ctrl+V."
                        : "Done. The picture guide is on your clipboard. Paste it with Ctrl+V.",
                    parts: parts.length,
                    copied: true
                }, id, parts);
            } else {
                finishDialog({
                    message: "The browser did not allow copying. Click Save picture to download it instead.",
                    parts: parts.length,
                    copied: false,
                    downloaded: parts.length,
                    copyError: copyError,
                    forceSave: true
                }, id, parts);
            }
        } catch (e) {
            last = { message: "Could not build the picture guide. " + ((e && e.message) || e), parts: 0 };
            window.EbobPicture.last = last;
            showDialog(last.message);
            var actions = dlg.querySelector(".eg-pic-actions");
            actions.innerHTML = "";
            var ok = document.createElement("button");
            ok.type = "button"; ok.className = "eg-btn eg-pic-cancel"; ok.textContent = "OK";
            ok.addEventListener("click", closeDlg); actions.appendChild(ok); ok.focus();
        }
    }

    var menu = null;
    function closeMenu() {
        if (menu) { menu.remove(); menu = null; }
        document.querySelectorAll("[data-share][aria-expanded=true]").forEach(function (b) { b.setAttribute("aria-expanded", "false"); });
    }
    function openMenu(btn) {
        closeMenu();
        var id = btn.getAttribute("data-share");
        opener = btn;
        btn.setAttribute("aria-expanded", "true");
        menu = document.createElement("div");
        menu.id = "egPicMenu";
        menu.className = "eg-ui eg-pic-menu";
        menu.setAttribute("role", "menu");
        menu.setAttribute("aria-label", "Picture guide");
        menu.innerHTML =
            '<p class="eg-pic-menu-kicker">Picture guide</p>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="copy">Copy PNG</button>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="download">Download PNG</button>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="email">Email guide</button>';
        document.body.appendChild(menu);
        var r = btn.getBoundingClientRect();
        var mw = menu.offsetWidth, mh = menu.offsetHeight;
        var left = Math.min(window.innerWidth - mw - 8, Math.max(8, r.right - mw));
        var top = r.bottom + 6;
        if (top + mh > window.innerHeight - 8) top = Math.max(8, r.top - mh - 6);
        menu.style.left = left + "px";
        menu.style.top = top + "px";
        var items = Array.prototype.slice.call(menu.querySelectorAll("[role=menuitem]"));
        items[0].focus();
        menu.addEventListener("click", function (e) {
            var it = e.target.closest("[data-act]");
            if (!it) return;
            run(it.getAttribute("data-act"), id);
        });
        menu.addEventListener("keydown", function (e) {
            var i = items.indexOf(document.activeElement);
            if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length].focus(); }
            else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
            else if (e.key === "Home") { e.preventDefault(); items[0].focus(); }
            else if (e.key === "End") { e.preventDefault(); items[items.length - 1].focus(); }
            else if (e.key === "Escape") { e.preventDefault(); closeMenu(); try { btn.focus(); } catch (err) {} }
            else if (e.key === "Enter" || e.key === " ") {
                var it = document.activeElement.closest("[data-act]");
                if (it) { e.preventDefault(); run(it.getAttribute("data-act"), id); }
            }
        });
    }

    function currentLesson() {
        var sp = qs();
        if (sp.get("lesson")) return LS.byId(sp.get("lesson"));
        var sc = sp.get("tutorial");
        if (sc && sc !== "free-mode") {
            for (var i = 0; i < LS.lessons.length; i++) if (LS.lessons[i].scenario === sc) return LS.lessons[i];
        }
        return null;
    }
    function addTopButton() {
        if (document.documentElement.classList.contains("eg-home")) return;
        if (inFrame()) return;
        var bar = document.querySelector(".eg-toolbar");
        if (!bar) { setTimeout(addTopButton, 60); return; }
        if (bar.querySelector("[data-share]")) return;
        var lesson = currentLesson();
        if (!lesson) return;
        var b = document.createElement("button");
        b.type = "button";
        b.className = "eg-tb-btn eg-share-btn";
        b.setAttribute("data-share", lesson.id);
        b.setAttribute("aria-haspopup", "menu");
        b.setAttribute("aria-expanded", "false");
        b.setAttribute("aria-label", "Share this lesson as a picture guide");
        b.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><path d="M4 12v7a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-7"/><path d="M8 8l4-4 4 4"/><path d="M12 4v12"/></svg><span>Share</span>';
        bar.insertBefore(b, bar.firstChild);
    }

    document.addEventListener("click", function (e) {
        var b = e.target.closest && e.target.closest("[data-share]");
        if (b && !e.target.closest("#egPicMenu")) {
            e.preventDefault();
            e.stopPropagation();
            if (menu && opener === b) closeMenu();
            else openMenu(b);
            return;
        }
        if (menu && !e.target.closest("#egPicMenu")) closeMenu();
    });
    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape" && menu) { closeMenu(); if (opener) try { opener.focus(); } catch (err) {} }
    });

    if (inFrame()) {
        document.documentElement.classList.add("eg-picture-mode");
        window.addEventListener("message", function (e) {
            if (!e.data || e.data.type !== "ebob-pic-run") return;
            driveSelf(function (t) { try { parent.postMessage({ type: "ebob-pic-progress", text: t }, "*"); } catch (err) {} })
                .then(function (steps) { parent.postMessage({ type: "ebob-pic-done", steps: steps }, "*"); })
                .catch(function (err) { parent.postMessage({ type: "ebob-pic-done", err: String(err && err.message || err) }, "*"); });
        });
    } else {
        if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", addTopButton);
        else addTopButton();
    }

    window.EbobPicture = {
        build: build,
        mailtoUrl: mailtoUrl,
        driveSelf: driveSelf,
        openMail: openMail,
        last: null
    };
})();
