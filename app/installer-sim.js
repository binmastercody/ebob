/*
 * Practice copy of the .NET 8 Binventory Installer (MainForm.cs / MainForm.Designer.cs).
 * Page numbers, control names, positions, colors and every label and message are taken from that code.
 * Nothing is installed: steps are timed, and the "Host PC", sites and DAT files are sample data.
 *   installer.html?lesson=<id>   mode comes from the lesson (fresh | upgrade | repair)
 *   installer.html?mode=upgrade  practice without a lesson
 */
(function () {
    "use strict";
    var TXT = window.EbobInstallerText || { license: "", siteDesc: "" };
    var sp = new URLSearchParams(location.search);
    var lesson = window.EbobLessons && sp.get("lesson") ? window.EbobLessons.byId(sp.get("lesson")) : null;
    var MODE = (lesson && lesson.installer && lesson.installer.mode) || sp.get("mode") || "fresh";
    var HOST_IP = "192.168.1.20";                       /* the sample Host PC that answers Connect */
    var EXT_SQL_IP = "192.168.1.30";                    /* a sample external SQL Server on its default instance */
    var EXISTING = MODE === "upgrade" ? "5.4.2" : MODE === "repair" ? "5.4.9" : "";
    var SITES = ["Current Location", "North Plant"];
    var DATS = [
        { file: "eBobExport_20261030_141522.dat", loc: "C:\\Users\\Public\\Documents\\Binventory Backups", date: "Oct 30, 2026", ver: "5.4.9" },
        { file: "eBobExport_20260612_090311.dat", loc: "C:\\Users\\Operator\\Documents\\Binventory Backups", date: "Jun 12, 2026", ver: "5.4.8" }
    ];
    var STEP_MS = 650, SQL_MS = 2200, CONNECT_MS = 2400;

    window.EbobGuideHost = {
        shellId: "instShell",
        waitText: "Waiting for the installer",
        refit: function () { if (window.__instReady) fit(); },
        ready: function () { return !!window.__instReady; }
    };

    function $(id) { return document.getElementById(id); }
    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
    function nl(s) { return esc(s).replace(/\n/g, "<br>"); }
    function px(x, y, w, h) { return "left:" + x + "px;top:" + y + "px;" + (w ? "width:" + w + "px;" : "") + (h ? "height:" + h + "px;" : ""); }
    function radio(id, name, x, y, text, extra) { return '<label class="iw-rad" id="' + id + '" style="' + px(x, y) + (extra || "") + '"><input type="radio" name="' + name + '"><span>' + esc(text) + "</span></label>"; }
    function check(id, x, y, text, cls) { return '<label class="iw-chk ' + (cls || "") + '" id="' + id + '" style="' + px(x, y) + '"><input type="checkbox"><span>' + esc(text) + "</span></label>"; }

    /* ---------------- window ---------------- */
    var desk = document.createElement("main");
    desk.id = "instDesk";
    desk.innerHTML =
        '<p class="inst-note">Practice copy of the Binventory Installer. Nothing is installed on this computer.</p>' +
        '<div class="iw" id="instShell" role="dialog" aria-label="Binventory Installer">' +
          '<div class="iw-titlebar"><img src="assets/eBob-V5-Application.ico" alt=""><span id="iwCaption">Binventory Installer</span><span class="iw-x" aria-hidden="true">&#x2715;</span></div>' +
          '<div class="iw-client" id="pnlContent">' +
            '<div class="iw-lbl" id="lblTitle" style="' + px(25, 20) + '">Binventory Installer v5.5</div>' +
            '<div class="iw-lbl" id="lblDescription" style="' + px(29, 56, 360) + '"></div>' +
            '<img id="picLogo" src="assets/installer/BinMaster-logo-200x-transparent.png" alt="BinMaster" style="' + px(400, 18, 155, 60) + '">' +

            /* page 0: installation type */
            '<div class="iw-pnl" id="pnlInstallType" style="' + px(33, 95, 530, 220) + '">' +
              '<fieldset class="iw-grp" id="grpInstallType" style="' + px(0, 15, 490, 125) + '"><legend>Installation Type</legend>' +
                radio("radServer", "itype", 20, 30, "Server/Host Install") +
                radio("radClient", "itype", 20, 58, "Client Install") +
                radio("radCustom", "itype", 20, 86, "Advanced") +
                radio("radRepairOption", "itype", 20, 86, "Repair Existing Installation", "display:none;") +
                '<div id="lblRepairError" style="' + px(40, 111) + 'display:none;">Repair is not available for versions below 5.4.9. Please use Remove instead.</div>' +
                '<div class="iw-opts" id="pnlCustomOptions" style="' + px(320, 10, 170, 113) + '">' +
                  check("chkWorkstation", 10, 12, "Workstation Client") + check("chkEngineScheduler", 10, 40, "eBob Services") + check("chkSqlServer", 10, 68, "SQL Server") +
                "</div>" +
              "</fieldset>" +
            "</div>" +
            '<div class="iw-pnl" id="pnlServerIP" style="' + px(33, 250, 530, 110) + '">' +
              '<label class="iw-lbl iw-b10" id="lblServerIP" for="txtServerIP" style="' + px(0, 10) + '">IP Address from Host PC:</label>' +
              '<input class="iw-txt" id="txtServerIP" type="text" spellcheck="false" autocomplete="off" style="' + px(0, 40, 250, 27) + '">' +
              '<button class="iw-btn iw-blue" id="btnTestConnection" type="button" style="' + px(260, 40, 120, 32) + '">Connect</button>' +
              '<div class="iw-lbl" id="lblConnectionProgress" style="' + px(260, 77, 250) + '">Connecting... Please Wait 30-60 Seconds</div>' +
              '<button class="iw-btn iw-sys" id="btnChangeIP" type="button" style="' + px(0, 72, 80, 23) + '">Change IP?</button>' +
            "</div>" +

            /* page 1: license */
            '<div class="iw-pnl" id="pnlLicense" style="' + px(25, 100, 530, 260) + '">' +
              '<textarea class="iw-txt" id="txtLicense" readonly style="' + px(0, 0, 530, 210) + '" aria-label="License agreement"></textarea>' +
              check("chkAcceptLicense", 0, 225, "I accept the terms of the license", "iw-c10") +
            "</div>" +

            /* page 15: site setup */
            '<div class="iw-pnl" id="pnlSiteSetup" style="' + px(25, 100, 530, 200) + '">' +
              '<fieldset class="iw-grp" id="grpSiteSetup" style="' + px(0, 0, 530, 200) + '"><legend>Site Setup</legend>' +
                '<label class="iw-lbl" id="lblEngineSite" for="cmbEngineSite" style="' + px(20, 35) + '">Engine Service Site Assignment:</label>' +
                '<select class="iw-txt" id="cmbEngineSite" style="' + px(20, 60, 480, 28) + '"><option value=""></option>' + SITES.map(function (s) { return "<option>" + esc(s) + "</option>"; }).join("") + "</select>" +
                '<div class="iw-lbl" id="lblEngineSiteDescription" style="' + px(20, 95, 480, 80) + '">' + nl(TXT.siteDesc) + "</div>" +
              "</fieldset>" +
            "</div>" +

            /* page 2: import project */
            '<div class="iw-pnl" id="pnlImportProject" style="' + px(25, 100, 530, 220) + '">' +
              '<fieldset class="iw-grp" id="grpImportProject" style="' + px(0, 0, 490, 90) + '"><legend>Project Data Import</legend>' +
                radio("radImportNow", "iproj", 20, 27, "Import Now") + radio("radImportLater", "iproj", 20, 55, "Import Later or New Setup") +
              "</fieldset>" +
              '<div class="iw-lbl iw-b10" id="lblDatFile" style="' + px(0, 105) + '">Project DAT File:</div>' +
              '<input class="iw-txt" id="txtDatFile" type="text" readonly style="' + px(0, 135, 380, 27) + '" aria-label="Project DAT File">' +
              '<button class="iw-btn iw-blue iw-b9" id="btnBrowseDat" type="button" style="' + px(390, 133, 90, 30) + '">Browse</button>' +
              '<button class="iw-btn iw-blue iw-b9" id="btnSearchDat" type="button" style="' + px(390, 168, 90, 30) + '">Search</button>' +
              '<div class="iw-lbl" id="lblDatStatus" style="' + px(0, 170, 380) + '">No file selected.</div>' +
            "</div>" +

            /* page 3: install path */
            '<div class="iw-pnl" id="pnlInstallPath" style="' + px(33, 100, 500, 110) + '">' +
              '<label class="iw-lbl iw-b10" id="lblInstallPath" for="txtInstallPath" style="' + px(0, 10) + '">Installation Directory:</label>' +
              '<input class="iw-txt" id="txtInstallPath" type="text" readonly value="C:\\Program Files (x86)\\eBob" style="' + px(0, 40, 500, 27) + '">' +
              '<div class="iw-lbl" id="lblDiskSpace" style="' + px(0, 75, 250) + '">Required Space: Calculating...</div>' +
              check("chkInstallUsbDrivers", 260, 75, "Install USB Drivers?") +
            "</div>" +

            /* pages 4 and 5: progress and completion */
            '<div class="iw-pnl" id="pnlProgress" style="' + px(25, 100, 530, 150) + '">' +
              '<div class="iw-bar" id="progressBar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" style="' + px(10, 30, 500, 32) + '"><i></i></div>' +
              '<div class="iw-lbl iw-r10" id="lblStatus" style="' + px(10, 75, 500) + '" aria-live="polite">Ready</div>' +
            "</div>" +

            /* page 6: upgrade type */
            '<div class="iw-pnl" id="pnlUpgradeType" style="' + px(33, 95, 530, 145) + '">' +
              '<fieldset class="iw-grp" id="grpUpgradeType" style="' + px(0, 15, 490, 125) + '"><legend>Upgrade Install Type</legend>' +
                radio("radUpgradeServer", "utype", 25, 40, "Server/HOST Install", "font-size:13.33px;") + radio("radUpgradeClient", "utype", 25, 80, "Client Install", "font-size:13.33px;") +
              "</fieldset>" +
            "</div>" +

            '<button class="iw-btn iw-white" id="btnCancel" type="button" style="' + px(25, 390, 100, 38) + '">Cancel</button>' +
            '<button class="iw-btn iw-white" id="btnBack" type="button" style="' + px(330, 390, 105, 38) + '">&lt; Back</button>' +
            '<button class="iw-btn iw-blue" id="btnNext" type="button" style="' + px(445, 390, 100, 38) + '">Next &gt;</button>' +
          "</div>" +
        "</div>" +
        '<div id="instClosed" class="inst-closed" hidden><p id="instClosedText"></p><button type="button" class="iw-btn iw-blue" id="instReopen">Open the installer again</button></div>' +
        '<div id="rebootScreen" class="inst-reboot" hidden><div class="inst-spin" aria-hidden="true"></div><p>Restarting</p></div>' +
        '<div id="mbLayer"></div>';
    document.body.appendChild(desk);
    $("txtLicense").value = TXT.license;

    /* ---------------- state ---------------- */
    var S = {
        page: 0, history: [], existing: !!EXISTING, installSql: true, installSvc: true,
        connected: false, connString: "", license: false, importNow: false, datValid: false, datPath: "", datVer: "",
        engineSite: "", busy: false, upgrade: false, op: "install", closed: false, resumed: false
    };
    window.__instState = S;

    function show(id, on) { var e = $(id); if (e) e.style.display = on ? "" : "none"; }
    function inputOf(id) { return $(id).querySelector("input"); }
    function checked(id) { return inputOf(id).checked; }
    function setChecked(id, v) { inputOf(id).checked = !!v; }
    function setText(id, t) { var e = $(id); var s = e.querySelector("span"); (s || e).textContent = t; }

    function selType() {
        if (checked("radServer")) return "server";
        if (checked("radClient")) return "client";
        if (checked("radCustom")) return "custom";
        if (checked("radRepairOption")) return "repair";
        return "";
    }
    function isRemoval() { return S.existing && selType() === "client"; }
    function isRepair() { return S.existing && (selType() === "repair" || (selType() === "server" && EXISTING === "5.4.9")); }
    function isUpgrade() { return S.upgrade; }
    function needsIp() { var t = selType(); return !S.existing && (t === "client" || (t === "custom" && !checked("chkSqlServer"))); }
    function computeParts() {
        var t = selType();
        if (t === "server") { S.installSql = true; S.installSvc = true; }
        else if (t === "client") { S.installSql = false; S.installSvc = false; }
        else if (t === "custom") { S.installSql = checked("chkSqlServer"); S.installSvc = checked("chkEngineScheduler"); }
    }
    function fmtMB(mb) { return mb >= 1024 ? (Math.round(mb / 1024 * 100) / 100) + " GB" : mb + " MB"; }
    function diskText() {
        var mb = 10 + 600 + 400 + (S.installSql ? 2048 + 1024 : 0);  /* app files, Crystal Reports and its temp, SQL Express and its temp */
        return "Required Space: " + fmtMB(mb) + " (Available: 212.4 GB)";
    }

    function resetConnection() {
        S.connected = false; S.connString = "";
        var t = $("txtServerIP"); t.disabled = false; t.classList.remove("iw-locked"); t.value = "";
        show("btnTestConnection", true); $("btnTestConnection").disabled = false;
        show("btnChangeIP", false); show("lblConnectionProgress", false);
    }

    /* ---------------- pages ---------------- */
    var PANELS = ["pnlInstallType", "pnlServerIP", "pnlLicense", "pnlSiteSetup", "pnlImportProject", "pnlInstallPath", "pnlProgress", "pnlUpgradeType"];
    function render() {
        var p = S.page;
        PANELS.forEach(function (id) { show(id, false); });
        var next = $("btnNext"), back = $("btnBack"), cancel = $("btnCancel"), d = $("lblDescription");
        next.disabled = false; back.disabled = false; cancel.disabled = false; next.textContent = "Next >";
        if (p === 0) {
            show("pnlInstallType", true);
            back.disabled = true;
            var t = selType();
            show("pnlCustomOptions", !S.existing && t === "custom");
            show("pnlServerIP", needsIp());
            if (S.existing) {
                d.textContent = "An existing Binventory installation has been detected. Choose to upgrade, repair, or remove it.";
                if (isRemoval()) next.textContent = "Uninstall"; else if (isRepair()) next.textContent = "Repair";
            } else d.textContent = "Select the type of installation you want to perform.";
        } else if (p === 1) {
            show("pnlLicense", true);
            d.textContent = "Please read and accept the license agreement to continue.";
            next.disabled = !S.license;
        } else if (p === 15) {
            show("pnlSiteSetup", true);
            d.textContent = "Select site assignments for Engine Service and Workstation.";
        } else if (p === 2) {
            show("pnlImportProject", true);
            d.textContent = "";
            ["txtDatFile", "btnBrowseDat", "btnSearchDat"].forEach(function (id) { $(id).disabled = !S.importNow; });
            $("lblDatFile").classList.toggle("iw-dim", !S.importNow);
            next.disabled = S.importNow && !S.datValid;
        } else if (p === 3) {
            show("pnlInstallPath", true);
            d.textContent = "Binventory Installation Path";
            show("chkInstallUsbDrivers", S.installSvc || S.installSql);
            $("lblDiskSpace").textContent = diskText();
            next.textContent = "Install";
        } else if (p === 4) {
            show("pnlProgress", true);
            d.textContent = S.op === "remove" ? "Removing existing Binventory installation..." : S.op === "repair" ? "Repairing Binventory installation..." : "Installing Binventory Inventory Management System...";
            back.disabled = true; next.disabled = true;
        } else if (p === 5) {
            show("pnlProgress", true);
            back.disabled = true; cancel.disabled = true;
            next.textContent = S.op === "remove" ? "Close" : "Finish";
            d.textContent = S.doneText || "Install Finished Successfully.";
        } else if (p === 6) {
            show("pnlUpgradeType", true);
            d.textContent = "Select the type of upgrade you want to perform.";
            show("pnlServerIP", checked("radUpgradeClient"));
        }
        document.documentElement.setAttribute("data-inst-page", String(p));
    }
    function go(p) { if (p !== S.page) S.history.push(S.page); S.page = p; render(); }

    /* ---------------- message boxes ---------------- */
    var ICONS = { info: "i", warn: "!", error: "\u2715", question: "?" };
    var Z = 1;
    function mount() { var l = $("mbLayer"); Array.prototype.forEach.call(l.children, function (c) { if (!c.classList.contains("mb-shade")) c.style.zoom = Z; }); }
    function msgBox(text, title, buttons, icon, cb) {
        var layer = $("mbLayer");
        var b = (buttons || "OK").split("|");
        layer.innerHTML = '<div class="mb-shade"></div><div class="mb" id="mbBox" role="alertdialog" aria-modal="true" aria-labelledby="mbTitle" aria-describedby="mbText">' +
            '<div class="mb-title" id="mbTitle">' + esc(title) + '<span class="iw-x" aria-hidden="true">&#x2715;</span></div>' +
            '<div class="mb-body"><span class="mb-ico mb-ico--' + (icon || "info") + '" aria-hidden="true">' + ICONS[icon || "info"] + '</span><div class="mb-text" id="mbText">' + nl(text) + "</div></div>" +
            '<div class="mb-foot">' + b.map(function (x) { return '<button type="button" class="mb-btn" id="mb' + x + '">' + x + "</button>"; }).join("") + "</div></div>";
        mount();
        b.forEach(function (x) { $("mb" + x).addEventListener("click", function () { layer.innerHTML = ""; if (cb) cb(x); }); });
        setTimeout(function () { var f = $("mb" + b[0]); if (f) f.focus(); }, 30);
    }
    function failDialog(cb) {
        var layer = $("mbLayer");
        layer.innerHTML = '<div class="mb-shade"></div><div class="cf" id="cfDialog" role="alertdialog" aria-modal="true" aria-labelledby="cfTitle">' +
            '<div class="mb-title">Connection Failed<span class="iw-x" aria-hidden="true">&#x2715;</span></div>' +
            '<div class="cf-client">' +
              '<div class="cf-ico" aria-hidden="true">\u26A0</div>' +
              '<div class="cf-h" id="cfTitle">Connection failed. Let\'s fix that!</div>' +
              '<div class="cf-q">Is your <b>Host PC</b> online?</div>' +
              '<div class="cf-yes" id="cfYes">' + nl("YES?:\n\n1. Go to Host PC\n2. Windows search\n3. Find and open Command Prompt\n4. In Command Prompt window after flashing cursor line type, IPCONFIG\n5. Note the IPV4 address\n6. Hit Retry button on this window\n7. Enter the IPV4 address\n8. Click Retry and Type IP, click Connect.") + "</div>" +
              '<div class="cf-no" id="cfNo">' + nl("NO?:\n\nCheck that Host is plugged into power and network. Plug it in!\nThen, check if Binventory software is installed on the Host PC.\nIf not, install Binventory Host!") + "</div>" +
              '<button type="button" class="iw-btn iw-blue" id="cfRetry">Retry</button>' +
            "</div></div>";
        mount();
        $("cfRetry").addEventListener("click", function () { layer.innerHTML = ""; if (cb) cb(); });
        setTimeout(function () { $("cfRetry").focus(); }, 30);
    }
    function picker(cb) {
        var layer = $("mbLayer");
        layer.innerHTML = '<div class="mb-shade"></div><div class="dp" id="dpDialog" role="dialog" aria-modal="true" aria-labelledby="dpTitle">' +
            '<div class="mb-title">Select Your Project File<span class="iw-x" aria-hidden="true">&#x2715;</span></div>' +
            '<div class="dp-client"><div class="dp-h" id="dpTitle">We found more than one project file</div>' +
            '<div class="dp-info">Pick the one you want to import, then click \u201CUse This File\u201D. The newest is already selected.</div>' +
            '<table class="dp-list" id="dpList"><thead><tr><th>File</th><th>Location</th><th>Date</th><th>Version</th></tr></thead><tbody>' +
            DATS.map(function (f, i) { return '<tr data-i="' + i + '"' + (i === 0 ? ' class="on"' : "") + "><td>" + esc(f.file) + "</td><td>" + esc(f.loc) + "</td><td>" + esc(f.date) + "</td><td>" + esc(f.ver) + "</td></tr>"; }).join("") +
            "</tbody></table>" +
            '<div class="dp-foot"><button type="button" class="iw-btn iw-blue" id="dpUse">Use This File</button><button type="button" class="iw-btn iw-white" id="dpCancel">Cancel</button></div></div></div>';
        mount();
        var sel = 0;
        Array.prototype.forEach.call(layer.querySelectorAll("#dpList tbody tr"), function (tr) {
            tr.addEventListener("click", function () { sel = +tr.getAttribute("data-i"); layer.querySelectorAll("#dpList tbody tr").forEach(function (r) { r.classList.toggle("on", r === tr); }); });
        });
        $("dpUse").addEventListener("click", function () { layer.innerHTML = ""; cb(DATS[sel]); });
        $("dpCancel").addEventListener("click", function () { layer.innerHTML = ""; cb(null); });
    }
    function browseDialog(cb) {
        /* stands in for the Windows "Select Project DAT File" open dialog */
        var layer = $("mbLayer");
        layer.innerHTML = '<div class="mb-shade"></div><div class="dp" id="ofDialog" role="dialog" aria-modal="true" aria-label="Select Project DAT File">' +
            '<div class="mb-title">Select Project DAT File<span class="iw-x" aria-hidden="true">&#x2715;</span></div>' +
            '<div class="dp-client"><div class="dp-info">' + esc(DATS[0].loc) + '</div><table class="dp-list"><thead><tr><th>Name</th><th>Date</th></tr></thead><tbody>' +
            '<tr class="on" id="ofFile"><td>' + esc(DATS[0].file) + "</td><td>" + esc(DATS[0].date) + "</td></tr></tbody></table>" +
            '<div class="dp-info">Files of type: DAT Files (*.dat)</div>' +
            '<div class="dp-foot"><button type="button" class="iw-btn iw-blue" id="ofOpen">Open</button><button type="button" class="iw-btn iw-white" id="ofCancel">Cancel</button></div></div></div>';
        mount();
        $("ofOpen").addEventListener("click", function () { layer.innerHTML = ""; cb(DATS[0]); });
        $("ofCancel").addEventListener("click", function () { layer.innerHTML = ""; cb(null); });
    }
    function useDat(f) {
        if (!f) return;
        S.datPath = f.loc + "\\" + f.file; S.datVer = f.ver; S.datValid = true;
        $("txtDatFile").value = S.datPath;
        $("lblDatStatus").textContent = "Project Validated. Source Version: " + f.ver;
        $("lblDatStatus").className = "iw-lbl iw-ok";
        render();
    }

    /* ---------------- progress ---------------- */
    function stepsFor() {
        var s = [];
        if (S.op === "repair") return ["Preparing Repair", "Restoring Missing Application Files", "Ensuring Configuration Files", "Ensuring Windows Services", "Recreating Shortcuts", "Finalizing Repair"];
        if (S.op === "remove") return ["Exporting Data Backup", "Killing eBob Processes", "Stopping eBob Services", "Removing eBob Services", "Removing SQL Server Instance (BINVENTORY)", "Removing Firewall Rules", "Removing Shortcuts", "Cleaning Registry Entries", "Removing Application Files"];
        if (S.op === "upgrade") return ["Exporting Data Backup", "Killing eBob Processes", "Stopping eBob Services", "Removing eBob Services", "Removing x86 Components", "Removing Application Files", "Cleaning Old Shortcuts", "Cleaning Registry Entries", "Removing Legacy Firewall Rules"];
        if (S.op === "upgrade-resume") return ["Extracting Application Files", "Installing USB Drivers", "Configuring Firewall Rules", "Validating Crystal Reports Runtime", "Installing Windows Services", "Ensuring Configuration Files", "Updating Database Configuration", "Recreating Shortcuts", "Restarting Services (Cycle 1)", "Refreshing Workstation (Cycle 1)", "Restarting Services (Cycle 2)", "Refreshing Workstation (Cycle 2)", "Finalizing Upgrade"];
        s.push("Extracting Application Files");
        if (checked("chkInstallUsbDrivers") && (S.installSql || S.installSvc)) s.push("Installing USB Drivers");
        if (S.installSql) s.push("Installing SQL Server Express");
        s.push("Installing Crystal Reports", "Installing Core Application Files", "Installing Configuration Files");
        if (S.installSvc) s.push("Installing Windows Services");
        if (S.installSql) { s.push("Creating Database"); if (S.importNow) s.push("Importing Project Data", "Selecting Site for Engine Service"); }
        if (S.installSvc || S.installSql) s.push("Configuring Firewall Rules");
        s.push("Creating Configuration Files", "Updating Database Configuration");
        if (S.installSvc) s.push("Restarting Services");
        s.push("Refreshing Workstation", "Creating Start Menu Shortcuts", "Finalizing Installation");
        return s;
    }
    var runTimer = 0;
    function setBar(pct) { var b = $("progressBar"); b.firstChild.style.width = pct + "%"; b.setAttribute("aria-valuenow", String(pct)); }
    function run() {
        var list = stepsFor(), i = 0;
        S.busy = true; S.steps = list; setBar(0); $("lblStatus").textContent = "Ready";
        go(4);
        (function nextStep() {
            if (S.closed) return;
            if (i >= list.length) { S.busy = false; return complete(); }
            var name = list[i];
            setBar(Math.floor(i * 100 / list.length)); $("lblStatus").textContent = name;
            runTimer = setTimeout(function () {
                $("lblStatus").textContent = "Completed: " + name;
                i++; setBar(Math.floor(i * 100 / list.length));
                runTimer = setTimeout(nextStep, 150);
            }, /SQL Server Express/.test(name) ? SQL_MS : STEP_MS);
        })();
    }
    function complete() {
        setBar(100);
        var op = S.op;
        if (op === "remove") {
            S.doneText = "Uninstall completed successfully.\n\nA restart is required to complete the uninstall.";
            $("lblStatus").textContent = "Uninstall completed successfully.";
            go(5); $("lblDescription").innerHTML = nl(S.doneText);
            msgBox("A restart is required to complete the uninstall.\n\nWould you like to restart now?", "Restart Required", "Yes|No", "question", function (a) {
                if (a === "Yes") reboot(function () { closeInstaller("The PC restarted and Binventory was removed."); });
                else msgBox("Please restart your computer to complete the uninstall.", "Restart Later", "OK", "info");
            });
            return;
        }
        if (op === "upgrade") {
            S.doneText = "Upgrade completed successfully.\n\nA restart is required to continue/finish the upgrade.";
            $("lblStatus").textContent = "Upgrade completed successfully.";
            go(5); $("lblDescription").innerHTML = nl(S.doneText);
            msgBox("A restart is required to finish the upgrade.\n\nWould you like to restart now?", "Restart Required", "Yes|No", "question", function (a) {
                if (a === "Yes") reboot(resumeUpgrade);
                else msgBox("The installer will continue upon next reboot.", "Restart Later", "OK", "info", function () {
                    closeInstaller("The installer will continue after the next restart.", "Restart the PC", function () { reboot(resumeUpgrade); });
                });
            });
            return;
        }
        S.doneText = op === "upgrade-resume" ? "Upgrade completed successfully." : op === "repair" ? "Repair completed successfully." : "Install Finished Successfully.";
        $("lblStatus").textContent = S.doneText;
        go(5);
    }
    function reboot(after) {
        $("instShell").hidden = true; $("instClosed").hidden = true; $("mbLayer").innerHTML = "";
        var r = $("rebootScreen"); r.hidden = false; S.rebooting = true;
        setTimeout(function () { r.hidden = true; S.rebooting = false; after(); }, 2200);
    }
    function resumeUpgrade() {
        /* RunOnce "BinventoryInstaller" starts the installer again with /resume=upgrade */
        S.closed = false; S.resumed = true; S.op = "upgrade-resume";
        $("instShell").hidden = false; $("instClosed").hidden = true;
        run();
    }
    /* Finish or Close: the window fades out (no fade with reduced motion) and is then removed from view.
       In a guided lesson the guide shows the end of the lesson, so no "installer closed" note is left behind. */
    function closeInstaller(text, btnText, btnAction, quiet) {
        clearTimeout(runTimer);
        S.closed = true; S.busy = false;
        $("mbLayer").innerHTML = "";
        var w = $("instShell"), reduce = document.documentElement.classList.contains("eg-reduce-motion");
        if (quiet) {
            if (reduce) { w.hidden = true; }
            else { w.classList.add("iw-closing"); setTimeout(function () { if (S.closed) { w.hidden = true; } w.classList.remove("iw-closing"); }, 180); }
            var note = document.querySelector(".inst-note"); if (note) note.hidden = true;
            return;
        }
        w.hidden = true;
        $("instClosedText").textContent = text;
        var b = $("instReopen");
        b.textContent = btnText || "Open the installer again";
        b.onclick = btnAction || function () { location.reload(); };
        $("instClosed").hidden = false;
    }

    /* ---------------- events ---------------- */
    function onType() {
        computeParts();
        if (needsIp() && !S.connected && !$("txtServerIP").value) resetConnection();
        render();
    }
    ["radServer", "radClient", "radCustom", "radRepairOption"].forEach(function (id) {
        inputOf(id).addEventListener("change", function () {
            if (id === "radClient" && !S.existing) resetConnection();
            if (id === "radCustom") { setChecked("chkEngineScheduler", false); setChecked("chkSqlServer", false); resetConnection(); }
            onType();
        });
    });
    ["chkEngineScheduler", "chkSqlServer"].forEach(function (id) { inputOf(id).addEventListener("change", onType); });
    inputOf("radUpgradeServer").addEventListener("change", render);
    inputOf("radUpgradeClient").addEventListener("change", function () { resetConnection(); render(); });
    inputOf("chkAcceptLicense").addEventListener("change", function () { S.license = checked("chkAcceptLicense"); render(); });
    inputOf("radImportNow").addEventListener("change", function () {
        S.importNow = true;
        if (!S.datPath) { S.datValid = false; $("lblDatStatus").textContent = "Select or drag a DAT file to continue."; $("lblDatStatus").className = "iw-lbl"; }
        render();
    });
    inputOf("radImportLater").addEventListener("change", function () { S.importNow = false; render(); });
    $("cmbEngineSite").addEventListener("change", function () { S.engineSite = this.value; });

    $("btnTestConnection").addEventListener("click", function () {
        var ip = $("txtServerIP").value.trim();
        if (!ip) return msgBox("Please enter a server IP address. The IP field cannot be blank.", "IP Address Required", "OK", "warn");
        if (!/^(25[0-5]|2[0-4]\d|1?\d?\d)(\.(25[0-5]|2[0-4]\d|1?\d?\d)){3}$/.test(ip)) return msgBox("Please enter a valid IP address (e.g., 192.168.1.100).", "Invalid IP Address", "OK", "warn");
        var btn = this; btn.disabled = true; show("lblConnectionProgress", true); S.connecting = true;
        setTimeout(function () {
            S.connecting = false; show("lblConnectionProgress", false);
            /* the installer tries ip\\BINVENTORY, then ip\\EBOB, then the default instance (ip) */
            if (ip === HOST_IP || ip === EXT_SQL_IP) {
                S.connected = true; S.connString = ip === HOST_IP ? ip + "\\BINVENTORY" : ip;
                show("btnTestConnection", false);
                var t = $("txtServerIP"); t.value = S.connString; t.disabled = true; t.classList.add("iw-locked");
                show("btnChangeIP", true);
                msgBox("IP Verified!\n\nSuccessfully connected to the database and found the system data.\n\nClick Next to proceed with the installation.", "Connection Successful", "OK", "info");
            } else {
                failDialog(function () { btn.disabled = false; });
            }
        }, CONNECT_MS);
    });
    $("btnChangeIP").addEventListener("click", function () { resetConnection(); render(); $("txtServerIP").focus(); });

    $("btnBrowseDat").addEventListener("click", function () { browseDialog(useDat); });
    $("btnSearchDat").addEventListener("click", function () {
        var b = this; b.textContent = "Searching..."; b.disabled = true;
        $("lblDatStatus").textContent = "Searching your computer for project files..."; $("lblDatStatus").className = "iw-lbl";
        setTimeout(function () { b.textContent = "Search"; b.disabled = false; picker(function (f) { if (f) useDat(f); else $("lblDatStatus").textContent = "Select or drag a DAT file to continue."; }); }, 1400);
    });
    var pImp = $("pnlImportProject");
    pImp.addEventListener("dragover", function (e) { if (S.importNow) e.preventDefault(); });
    pImp.addEventListener("drop", function (e) {
        e.preventDefault();
        var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (!f) return;
        if (!/\.dat$/i.test(f.name)) return msgBox("Please drop a .dat project file.", "Invalid File", "OK", "warn");
        useDat({ file: f.name, loc: "C:\\Users\\Operator\\Desktop", date: "", ver: "5.4.9" });
    });

    $("btnCancel").addEventListener("click", function () {
        if (S.busy) {
            msgBox("Installation is in progress. Are you sure you want to cancel?", "Cancel Installation", "Yes|No", "question", function (a) {
                if (a === "Yes") closeInstaller("Installation failed. All changes have been rolled back.");
            });
            return;
        }
        closeInstaller("The installer closed without making changes.");
    });
    $("btnBack").addEventListener("click", function () {
        if (!S.history.length || S.page === 4 || S.page === 5) return;
        S.page = S.history.pop(); render();
    });
    $("btnNext").addEventListener("click", function () {
        var p = S.page;
        if (p === 0) {
            computeParts();
            if (S.existing) {
                if (selType() === "server" && EXISTING !== "5.4.9") { S.upgrade = true; return go(6); }
                if (isRemoval()) { S.op = "remove"; return run(); }
                if (isRepair()) { S.op = "repair"; return run(); }
            }
            if (needsIp()) {
                if (!$("txtServerIP").value.trim()) return msgBox("Please enter a server IP address.", "IP Address Required", "OK", "warn");
                if (!S.connected) return msgBox("Please click Connect to verify connectivity before proceeding.", "Connection Test Required", "OK", "warn");
            }
            return go(1);
        }
        if (p === 1) {
            if (!S.license) return;
            if (selType() === "custom" && checked("chkEngineScheduler") && S.connected) return go(15);
            return go(S.installSql ? 2 : 3);
        }
        if (p === 15) {
            if (!$("cmbEngineSite").value) return msgBox("Please select a site for Engine Service assignment before proceeding.", "Site Selection Required", "OK", "warn");
            return go(S.installSql ? 2 : 3);
        }
        if (p === 2) {
            if (S.importNow && !S.datValid) return msgBox("Please select a valid project DAT file before continuing.", "DAT File Required", "OK", "warn");
            return go(3);
        }
        if (p === 3) { S.op = "install"; return run(); }
        if (p === 5) { return closeInstaller(S.op === "remove" ? "The installer closed." : "The installer closed. Binventory Workstation starts on its own after a successful install, upgrade or repair.", null, null, !!lesson); }
        if (p === 6) {
            if (checked("radUpgradeClient") && !S.connected) return msgBox("Please click Connect to verify connectivity before proceeding.", "Connection Test Required", "OK", "warn");
            S.op = "upgrade"; return run();
        }
    });

    /* ---------------- start ---------------- */
    setChecked("chkWorkstation", true); inputOf("chkWorkstation").disabled = true;
    setChecked("chkInstallUsbDrivers", true);
    setChecked("radImportLater", true);
    setChecked("radUpgradeServer", true);
    show("lblConnectionProgress", false); show("btnChangeIP", false);
    $("txtServerIP").value = "192.168.1.100";
    if (S.existing) {
        setText("radServer", EXISTING === "5.4.9" ? "Repair Existing Installation" : "Upgrade Existing Installation");
        setText("radClient", "Remove Existing Binventory Installation");
        show("radCustom", false);
        if (EXISTING !== "5.4.9") {
            show("radRepairOption", true); inputOf("radRepairOption").disabled = true; $("radRepairOption").classList.add("iw-off");
            show("lblRepairError", true);
            $("grpInstallType").style.height = (111 + 13 + 15) + "px";
        }
        setChecked("radClient", true);
    } else setChecked("radServer", true);
    computeParts();
    render();

    function fit() {
        var top = (document.querySelector(".eg-topbar") || { offsetHeight: 52 }).offsetHeight;
        /* composed: the guide card sits in a reserved slot left of the window, so leave room for it */
        var comp = document.documentElement.classList.contains("eg-composed");
        var dockW = comp ? parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--eg-dock-w")) || 0 : 0;
        var zw = comp ? (window.innerWidth - 40 - dockW) / 560 : (window.innerWidth * 0.6) / 560;
        var z = Math.min((window.innerHeight - top - 40) / 482, zw, 1.45);
        z = Math.max(0.8, Math.round(z * 100) / 100);
        Z = z;
        $("instShell").style.zoom = z;
        /* the margin is inside the zoom, so divide to get dockW screen pixels */
        $("instShell").style.marginLeft = comp ? (dockW / z) + "px" : "0px";
        var w = $("instShell"), l = $("mbLayer");
        if (!w.hidden) {
            var r = w.getBoundingClientRect(), d = desk.getBoundingClientRect();
            l.style.left = (r.left - d.left) + "px"; l.style.top = (r.top - d.top) + "px"; l.style.width = r.width + "px"; l.style.height = r.height + "px";
        }
        mount();
        /* hide the practice note when the window or the guide card would sit on it */
        var note = document.querySelector(".inst-note");
        if (note) {
            note.style.visibility = "";
            var nr = note.getBoundingClientRect(), hit = false;
            [w, document.querySelector(".eg-coach")].forEach(function (e) {
                if (!e || e.hidden) return; var q = e.getBoundingClientRect();
                if (q.width && nr.left < q.right && nr.right > q.left && nr.top < q.bottom && nr.bottom > q.top) hit = true;
            });
            if (hit) note.style.visibility = "hidden";
        }
    }
    window.addEventListener("resize", fit);
    setTimeout(function () {
        fit();
        window.__instReady = true;
        if (S.existing) msgBox("Binventory install detected.\n\nVersion: " + EXISTING + "\nLocation: C:\\Program Files (x86)\\eBob\n\nPlease review the installation options on the next screen.", "Binventory Install Detected", "OK", "info");
    }, 250);

    window.EbobInstaller = { state: S, page: function () { return S.page; }, open: function () { return !S.closed && !S.rebooting; }, dialog: function () { return !!$("mbLayer").firstChild; } };
})();
