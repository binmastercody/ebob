/* Share a lesson as its PDF guide: Open PDF, Download PDF, Email PDF, Copy PDF link.
   Every lesson has a PDF in pdf/lessons/<id>.pdf, built by tools/build-pdfs.mjs in the same
   layout as the Binventory 1.0 Vessel Configuration PDF. Email PDF saves the PDF to Downloads
   and opens a new email with the cursor at the top, ready to attach it. */
(function () {
    "use strict";
    var LS = window.EbobLessons;
    if (!LS) return;
    var SITE = "https://binmastercody.github.io/ebob/";
    var qs = new URLSearchParams(location.search);
    if (qs.get("pdf") === "1") return;

    function hrefOf(l) {
        return l.page === "installer" ? "installer.html?lesson=" + l.id
            : l.kind === "doc" ? "doc.html?lesson=" + l.id
            : l.scenario ? "emulator.html?tutorial=" + l.scenario + "&fresh=1"
            : "emulator.html?tutorial=free-mode&lesson=" + l.id;
    }
    function pdfHref(l) { return "pdf/lessons/" + l.id + ".pdf"; }
    function pdfName(l) { return "Binventory 1.0 - " + String(l.title).replace(/[\\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim() + ".pdf"; }

    function clickLink(href, extra) {
        var a = document.createElement("a");
        a.className = "eg-ui";
        a.href = href;
        if (extra) for (var k in extra) a.setAttribute(k, extra[k]);
        a.style.cssText = "position:fixed;left:-9999px;top:0;width:1px;height:1px;";
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { a.remove(); }, 1000);
    }
    function download(l) { clickLink(pdfHref(l), { download: pdfName(l) }); }
    /* no instructions in the email: the cursor lands on the empty first line, then the links */
    function mailtoUrl(l) {
        var body = ["", "", l.title, "PDF guide: " + SITE + pdfHref(l), "Online lesson: " + SITE + hrefOf(l)].join("\r\n");
        return "mailto:?subject=" + encodeURIComponent("Binventory guide: " + l.title) + "&body=" + encodeURIComponent(body);
    }
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

    var opener = null, dlg = null, menu = null;
    function closeDlg() {
        if (dlg && dlg.parentNode) dlg.remove();
        dlg = null;
        if (opener) { try { opener.focus(); } catch (e) {} }
    }
    function showDialog(text, l) {
        if (dlg) dlg.remove();
        dlg = document.createElement("div");
        dlg.className = "eg-ui eg-pic-dlg";
        dlg.innerHTML = '<div class="eg-pic-card" role="dialog" aria-modal="true" aria-labelledby="egPicDlgTitle">' +
            '<h2 id="egPicDlgTitle" class="eg-pic-dlg-title">PDF guide</h2>' +
            '<p class="eg-pic-status" role="status"></p>' +
            '<div class="eg-pic-actions"><button type="button" class="eg-btn eg-pic-save">Save PDF again</button>' +
            '<button type="button" class="eg-btn eg-btn--ghost eg-pic-cancel">OK</button></div></div>';
        dlg.querySelector(".eg-pic-status").textContent = text;
        document.body.appendChild(dlg);
        var save = dlg.querySelector(".eg-pic-save"), ok = dlg.querySelector(".eg-pic-cancel");
        save.addEventListener("click", function () { download(l); });
        ok.addEventListener("click", closeDlg);
        dlg.querySelector(".eg-pic-card").addEventListener("keydown", function (e) {
            if (e.key === "Escape") { e.preventDefault(); closeDlg(); return; }
            if (e.key !== "Tab") return;
            e.preventDefault();
            (document.activeElement === ok ? save : ok).focus();
        });
        setTimeout(function () { ok.focus(); }, 0);
    }

    function run(act, id) {
        closeMenu();
        var l = LS.byId(id);
        if (!l) return;
        if (act === "open") { window.open(pdfHref(l), "_blank", "noopener"); return; }
        if (act === "download") { download(l); showDialog("Done. The PDF guide “" + pdfName(l) + "” is in your Downloads folder.", l); return; }
        if (act === "link") {
            copyText(SITE + pdfHref(l)).then(function () { showDialog("Done. The link to the PDF guide is copied. Paste it with Ctrl+V.", l); },
                function () { window.prompt("Copy this link:", SITE + pdfHref(l)); });
            return;
        }
        if (act === "email") {
            download(l);
            setTimeout(function () { clickLink(mailtoUrl(l)); }, 300);
            showDialog("Done. Your email opens now. The PDF guide “" + pdfName(l) + "” is in your Downloads folder: drag it into the email, or click Attach and pick it.", l);
        }
    }

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
        menu.setAttribute("aria-label", "PDF guide");
        menu.innerHTML =
            '<p class="eg-pic-menu-kicker">PDF guide</p>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="open">Open PDF</button>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="download">Download PDF</button>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="email">Email PDF</button>' +
            '<button type="button" class="eg-pic-item" role="menuitem" data-act="link">Copy PDF link</button>';
        document.body.appendChild(menu);
        var r = btn.getBoundingClientRect();
        var mw = menu.offsetWidth, mh = menu.offsetHeight;
        menu.style.left = Math.min(window.innerWidth - mw - 8, Math.max(8, r.right - mw)) + "px";
        var top = r.bottom + 6;
        if (top + mh > window.innerHeight - 8) top = Math.max(8, r.top - mh - 6);
        menu.style.top = top + "px";
        var items = Array.prototype.slice.call(menu.querySelectorAll("[role=menuitem]"));
        items[0].focus();
        menu.addEventListener("click", function (e) {
            var it = e.target.closest("[data-act]");
            if (it) run(it.getAttribute("data-act"), id);
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
        if (qs.get("lesson")) return LS.byId(qs.get("lesson"));
        var sc = qs.get("tutorial");
        if (sc && sc !== "free-mode") for (var i = 0; i < LS.lessons.length; i++) if (LS.lessons[i].scenario === sc) return LS.lessons[i];
        return null;
    }
    function addTopButton() {
        if (document.documentElement.classList.contains("eg-home")) return;
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
        b.setAttribute("aria-label", "Share this lesson as a PDF guide");
        b.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><path d="M4 12v7a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-7"/><path d="M8 8l4-4 4 4"/><path d="M12 4v12"/></svg><span>PDF</span>';
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
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", addTopButton);
    else addTopButton();

    window.EbobShare = { pdfHref: pdfHref, pdfName: pdfName, mailtoUrl: mailtoUrl, run: run };
})();
