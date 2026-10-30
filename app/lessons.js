/*
 * eBob Guide lesson catalog. Used by the lesson menu (index.html) and the guide engine (emulator.html).
 * Step content follows the Binventory V5.4.9 Operator's Manual (925-0267 Rev M); section numbers are in `manual`.
 *
 * Step fields
 *   title, msg      text for the coach card (msg may hold <b>)
 *   target          CSS selector or function returning the element to circle
 *   do              click | type | select | check | uncheck | read | wait
 *   value           text to type, or option label to select
 *   until           optional function; for click steps the step is done when this is true after the click
 *   opensMenu       true for a step that opens a menu (the guide steps back here if the menu closes)
 */
(function () {
    "use strict";

    /* ---------- small DOM helpers used by lesson conditions ---------- */
    function el(t) {
        if (!t) return null;
        if (t.nodeType === 1) return t;
        return typeof t === "function" ? t() : document.querySelector(t);
    }
    function visible(t) {
        var e = el(t);
        if (!e || !e.isConnected) return false;
        var r = e.getBoundingClientRect();
        if (r.width < 1 || r.height < 1) return false;
        var cs = getComputedStyle(e);
        return cs.visibility !== "hidden" && cs.display !== "none";
    }
    /** true when an open dialog's caption contains the text */
    function win(text) {
        var caps = document.querySelectorAll(".backdrop.show .modal-cap-title");
        for (var i = 0; i < caps.length; i++) if (caps[i].textContent.indexOf(text) >= 0) return true;
        return false;
    }
    function shown(id) { var b = document.getElementById(id); return !!(b && b.classList.contains("show")); }
    function menuBtn(name) {
        var bs = document.querySelectorAll("#menuStrip .menu-root > button");
        for (var i = 0; i < bs.length; i++) if (bs[i].textContent.trim() === name) return bs[i];
        return null;
    }
    function menuOpen(name) { var b = menuBtn(name); return !!(b && b.parentElement.classList.contains("open")); }
    /** the footer Close button of a dialog backdrop (not the x in the caption) */
    function closeBtn(backdropId) {
        return function () {
            var b = document.getElementById(backdropId);
            if (!b) return null;
            var bs = b.querySelectorAll("button");
            for (var i = bs.length - 1; i >= 0; i--) if (bs[i].textContent.trim() === "Close" && visible(bs[i])) return bs[i];
            return null;
        };
    }
    /** the caption bar of the front-most open dialog */
    function topCaption() {
        var bs = document.querySelectorAll(".backdrop.show"), best = null, bz = -1;
        for (var i = 0; i < bs.length; i++) { var z = parseInt(getComputedStyle(bs[i]).zIndex, 10) || 0; if (z >= bz) { bz = z; best = bs[i]; } }
        return best && best.querySelector(".modal-cap");
    }
    function vesselTile(id) { return '#vesselGrid .vessel[data-vessel-id="' + id + '"]'; }
    function pageTab(n) { return '#tabStrip [data-tab="' + (n - 1) + '"]'; }
    function tabActive(n) { var t = document.querySelector(pageTab(n)); return !!(t && t.classList.contains("active")); }
    function vsTab(name) { return '#backdropAppStack [data-vs-tab="' + name + '"]'; }
    function vsTabShown(fieldSel) { return function () { return visible(fieldSel); }; }
    function textOf(sel) { var e = el(sel); return e ? e.textContent.trim() : ""; }
    function titleBar() {
        var c = document.querySelectorAll("#appShell .title, .app-shell .title");
        for (var i = 0; i < c.length; i++) if (/Site:/.test(c[i].textContent)) return c[i];
        return null;
    }

    /* ---------- step builders ---------- */
    function menu(name, itemSel, itemLabel, opensWin, extraMsg) {
        return [
            { title: "Open the " + name + " menu", msg: "Click <b>" + name + "</b> in the menu bar at the top of the Binventory window.",
              target: function () { return menuBtn(name); }, do: "click", opensMenu: true,
              until: function () { return menuOpen(name); } },
            { title: "Choose " + itemLabel, msg: "Click <b>" + itemLabel + "</b>." + (extraMsg ? " " + extraMsg : ""),
              target: "#menuStrip " + itemSel, do: "click",
              until: opensWin ? function () { return win(opensWin); } : null }
        ];
    }
    function flat(list) { return [].concat.apply([], list); }

    var SECTIONS = [
        { id: "setup", title: "Setup", kind: "path", icon: "rocket", acc: "#5b3fa6", tint: "#f0ecfa",
          blurb: "Plan, install and set up a new system, from checking the PC to the first reading, in the order the manual uses." },
        { id: "alerts", title: "Alerts and email", icon: "bell", acc: "#c2410c", tint: "#fff1e8",
          blurb: "Level alarms, email server, contacts and who gets which email." },
        { id: "readings", title: "Readings and schedules", icon: "gauge", acc: "#0f766e", tint: "#e6f6f4",
          blurb: "Take measurements, read vessel details, group vessels and schedule readings." },
        { id: "reports", title: "Reports", icon: "chart", acc: "#1d4ed8", tint: "#e9effd",
          blurb: "Inventory reports, history charts and the emailed Site Status Report." },
        { id: "admin", title: "Users, backup and upgrades", icon: "users", acc: "#4b5563", tint: "#f1f2f4",
          blurb: "Add users with the right access, back up your system and upgrade safely." },
        { id: "trouble", title: "Troubleshooting", icon: "wrench", acc: "#b91c1c", tint: "#fdecec",
          blurb: "Find the cause of common problems by checking the right screens, step by step." }
    ];

    var L = [];

    /* ======================= SETUP ======================= */
    L.push({ id: "tour", section: "setup", title: "First launch: tour the Workstation", manual: "5, 5.1, 5.2",
        desc: "The menu bar, vessel panels, status and pages.",
        outro: "You know where the menus are, what a vessel panel shows and how to switch pages.",
        steps: [
            { title: "This is Binventory Workstation", msg: "Every vessel and setting is reached from this window. This practice copy starts with a sample site of 32 vessels. Nothing you do here touches a real system.",
              target: "#appShell", do: "read", ring: false },
            { title: "The menu bar", msg: "Binventory is run from six menus: <b>File</b>, <b>Vessel</b>, <b>Reports</b>, <b>Specifications</b>, <b>Site Assignment</b> and <b>Help</b>. Most setup is under Specifications and Vessel.",
              target: "#menuStrip", do: "read" },
            { title: "One panel per vessel", msg: "Each panel shows the vessel name, its contents, a level graphic, and the product height, volume and weight.",
              target: vesselTile("v1"), do: "read" },
            { title: "Sensor status", msg: "<b>Status</b> tells you what the sensor is doing. <b>Unknown</b> means the sensor has not measured since the PC restarted or the Binventory Engine was restarted.",
              target: vesselTile("v1") + " .vessel-status-val", do: "read" },
            { title: "Headroom Display", msg: "Check <b>Headroom Display</b> to show the empty space in the vessel instead of the product level.",
              target: vesselTile("v1") + ' [data-vessel-action="headroom"]', do: "check" },
            { title: "Switch it back", msg: "Uncheck <b>Headroom Display</b> to show the product level again.",
              target: vesselTile("v1") + ' [data-vessel-action="headroom"]', do: "uncheck" },
            { title: "Vessels are on pages", msg: "Click the <b>Page 2</b> tab to see the next set of vessels.",
              target: pageTab(2), do: "click", until: function () { return tabActive(2); } },
            { title: "Back to Page 1", msg: "Click <b>Page 1</b>.",
              target: pageTab(1), do: "click", until: function () { return tabActive(1); } }
        ] });

    L.push({ id: "system-setup", section: "setup", title: "System Setup", manual: "5.4.1",
        desc: "Registered user, how long readings are kept and the login rule.",
        outro: "System Setup holds the registered user, how many days of readings are kept and whether users must log in.",
        steps: flat([
            menu("Specifications", '[data-action="system-setup"]', "System Setup", "System Setup"),
            [
            { title: "Registered user", msg: "Type <b>Jane Operator</b> as the Registered User.", target: "#ssRegUser", do: "type", value: "Jane Operator" },
            { title: "Company and address", msg: "These fields are optional information about the registered user. They are filled in for the sample company.", target: "#ssCompany", do: "read" },
            { title: "Measurement Retention Days", msg: "Readings older than this many days are permanently deleted. The default is 60, about 2 months. Type <b>90</b> to keep about 3 months.", target: "#ssRetention", do: "type", value: "90" },
            { title: "Login rule", msg: "<b>Do NOT require users to login</b> is checked by default. With it checked, anyone who can use this PC has administrator access. The manual recommends unchecking it once users are set up. Leave it for now.", target: "#ssAutoLogin", do: "read" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#ssSave", do: "click", until: function () { return !win("System Setup"); } }
            ]
        ]) });

    L.push({ id: "site-setup", section: "setup", title: "Site Setup", manual: "5.4.2",
        desc: "Name the site, check distance units and the Engine host.",
        outro: "The site now has a clear name. Distance units and the Engine host address are set in Site Setup.",
        steps: flat([
            menu("Specifications", '[data-action="site-maintenance"]', "Site Maintenance", "Site Maintenance"),
            [
            { title: "Pick the site", msg: "Click <b>Current Location</b> in the list to highlight it.", target: '#backdropApp tr[data-sid="st1"]', do: "click" },
            { title: "Open it", msg: "Click <b>Select</b>.", target: "#backdropApp #smBtnSelect", do: "click", until: function () { return win("Site Setup"); } },
            { title: "Name the site", msg: "Name/Description is required and should be unique. Type <b>Main Plant</b>.", target: "#ssuName", do: "type", value: "Main Plant" },
            { title: "Distance units", msg: "Feet is the default. These units are used for vessel dimensions and sensor settings.", target: "#ssuDist", do: "read" },
            { title: "Engine host", msg: "<b>127.0.0.1</b> is only valid for a single stand-alone system. On a networked system, type the address of the PC that runs the Engine and Scheduler services. The Host Port, 8093, should generally not be changed.", target: "#ssuHostIp", do: "read" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#ssuSave", do: "click", until: function () { return !win("Site Setup"); } },
            { title: "Close Site Maintenance", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("Site Maintenance"); } }
            ]
        ]) });

    L.push({ id: "sensor-network", section: "setup", title: "Adding a sensor network", manual: "5.4.5",
        desc: "Add a network, pick its protocol and port, then save.",
        outro: "You added a sensor network. Every vessel on it will use the protocol and port you set here.",
        steps: flat([
            menu("Specifications", '[data-action="sensor-networks"]', "Sensor Networks", "Sensor Networks"),
            [
            { title: "Add a network", msg: "Click <b>Add New</b>. A new network is added to the list and highlighted.", target: "#snAddNew", do: "click",
              until: function () { return document.querySelectorAll("#backdropApp tr[data-nid]").length > 1; } },
            { title: "Open its setup", msg: "With the new network highlighted, click <b>Select</b> to open Sensor Network Setup.", target: "#snBtnSelect", do: "click",
              until: function () { return shown("backdropSnSetup"); } },
            { title: "Name it", msg: "Use a unique, descriptive name. Type <b>Modbus Network</b>.", target: "#snsName", do: "type", value: "Modbus Network" },
            { title: "Choose the protocol", msg: "All sensors on a network use one protocol. Choose <b>Modbus/RTU</b>.", target: "#snsProtocol", do: "select", value: "Modbus/RTU" },
            { title: "Interface", msg: "This network uses the PC's local serial port <b>COM3</b>. A Remote Serial Server needs an IP address and port instead.", target: "#snsCom", do: "read" },
            { title: "Communication parameters", msg: "These now show the Modbus/RTU settings: 9,600 bps, 8 data bits, no parity, 1 stop bit. They must match your sensors.", target: "#snsBaud", do: "read" },
            { title: "Advanced settings", msg: "Only change the advanced settings under guidance of BinMaster technical support.", target: "#snsAtt", do: "read" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#snSetupSave", do: "click", until: function () { return !shown("backdropSnSetup"); } },
            { title: "Check the list", msg: "<b>Modbus Network</b> is now in the list.", target: function () {
                  var rows = document.querySelectorAll("#backdropApp tr[data-nid]");
                  for (var i = 0; i < rows.length; i++) if (/Modbus Network/.test(rows[i].textContent)) return rows[i];
                  return null; }, do: "read" },
            { title: "Close Sensor Networks", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("Sensor Networks"); } }
            ]
        ]) });

    L.push({ id: "add-vessel", section: "setup", title: "Adding a vessel", manual: "5.4.6, 5.4.7",
        desc: "Name, shape, contents and sensor for a new vessel.",
        outro: "Your vessel has a name, a shape, its contents and a sensor. It now shows as a panel on the main screen.",
        steps: flat([
            menu("Vessel", '[data-action="vessel-maintenance"]', "Vessel Maintenance", "Vessel Maintenance"),
            [
            { title: "Add a vessel", msg: "Click <b>Add New</b>. Vessel Setup opens on the <b>Display</b> tab.", target: "#vmBtnAddNew", do: "click", until: function () { return visible("#vs_name"); } },
            { title: "Vessel name", msg: "Type <b>Corn 1</b>. Names can be up to 10 characters.", target: "#vs_name", do: "type", value: "Corn 1" },
            { title: "Fill color", msg: "Choose <b>Goldenrod</b> as the fill color for the level graphic.", target: "#vs_color", do: "select", value: "Goldenrod" },
            { title: "Units", msg: "Pick the volume and weight units shown for this vessel. Cubic feet and pounds are fine.", target: "#vs_vol_units", do: "read" },
            { title: "Type/Shape tab", msg: "Click the <b>Type/Shape</b> tab.", target: vsTab("typeshape"), do: "click", until: vsTabShown("#vs_vessel_type") },
            { title: "Vessel type", msg: "Choose <b>Vertical Cylinder with Cone</b>.", target: "#vs_vessel_type", do: "select", value: "Vertical Cylinder with Cone" },
            { title: "Height", msg: "Dimensions are in the site's distance units, feet. Type <b>40</b> for Height (H).", target: "#vs_sp_0", do: "type", value: "40" },
            { title: "Width", msg: "Type <b>18</b> for Width (W).", target: "#vs_sp_1", do: "type", value: "18" },
            { title: "Cone height", msg: "Type <b>8</b> for Cone Height (CH).", target: "#vs_sp_2", do: "type", value: "8" },
            { title: "Outlet width", msg: "Type <b>1</b> for Outlet Width (OW).", target: "#vs_sp_3", do: "type", value: "1" },
            { title: "Contents tab", msg: "Click the <b>Contents</b> tab.", target: vsTab("contents"), do: "click", until: vsTabShown("#vs_contents") },
            { title: "Product", msg: "Type <b>Corn</b> as the contents.", target: "#vs_contents", do: "type", value: "Corn" },
            { title: "Density", msg: "Weight is worked out from density. Type <b>45</b>, in pounds per cubic foot.", target: "#vs_density_val", do: "type", value: "45" },
            { title: "Sensor tab", msg: "Click the <b>Sensor</b> tab.", target: vsTab("sensor"), do: "click", until: vsTabShown("#vs_network") },
            { title: "Sensor network", msg: "The vessel uses <b>COM2 Sensor Network</b>, the SmartBob network on this sample site. The sensor type list only shows sensors that work on that network.", target: "#vs_network", do: "read" },
            { title: "Sensor type", msg: "Choose <b>SmartBob-II</b>.", target: "#vs_sensor_type", do: "select", value: "SmartBob-II" },
            { title: "Enable the sensor", msg: "<b>Enabled</b> must be checked for a new vessel. Make sure it is checked.", target: "#vs_sb_en1", do: "check" },
            { title: "Sensor address", msg: "Each sensor needs its own address, and it must match the address set on the sensor. Type <b>33</b>.", target: "#vs_sb_addr1", do: "type", value: "33" },
            { title: "Save the vessel", msg: "Click <b>Save</b>.", target: "#vsSave", do: "click", until: function () { return !visible("#vs_name"); } },
            { title: "Close Vessel Maintenance", msg: "Click <b>Close</b>. The main screen updates with the new vessel.", target: "#vmBtnClose", do: "click", until: function () { return !win("Vessel Maintenance"); } },
            { title: "Find the new vessel", msg: "New vessels go at the end. Click the <b>Page 3</b> tab.", target: pageTab(3), do: "click", until: function () { return tabActive(3); } },
            { title: "Your vessel", msg: "<b>Corn 1</b> has its own panel. If its sensor is wired and powered, click Measure to confirm it reads.", target: function () {
                  var t = document.querySelectorAll("#vesselGrid .vessel");
                  for (var i = 0; i < t.length; i++) if (/Corn 1/.test(t[i].textContent) && visible(t[i])) return t[i];
                  return null; }, do: "read" }
            ]
        ]) });

    /* ======================= ALERTS ======================= */
    function openVessel(vid, name) {
        return flat([
            menu("Vessel", '[data-action="vessel-maintenance"]', "Vessel Maintenance", "Vessel Maintenance"),
            [
            { title: "Pick a vessel", msg: "Click <b>" + name + "</b> in the list.", target: '#backdropApp tr[data-vid="' + vid + '"]', do: "click" },
            { title: "Open Vessel Setup", msg: "Click <b>Select</b>.", target: "#vmBtnSelect", do: "click", until: function () { return visible("#vs_name"); } }
            ]
        ]);
    }
    var saveAndCloseVessel = [
        { title: "Save", msg: "Click <b>Save</b>.", target: "#vsSave", do: "click", until: function () { return !visible("#vs_name"); } },
        { title: "Close Vessel Maintenance", msg: "Click <b>Close</b>.", target: "#vmBtnClose", do: "click", until: function () { return !win("Vessel Maintenance"); } }
    ];

    L.push({ id: "level-alarms", section: "alerts", title: "Setting level alarms", manual: "5.4.7.5",
        desc: "Turn on pre-high and pre-low warnings for a vessel.",
        outro: "When a reading passes an enabled alarm level, the vessel shows a blinking red alarm on the main screen.",
        steps: flat([
            openVessel("v2", "Cem Silo 2"),
            [
            { title: "Alarms tab", msg: "Click the <b>Alarms</b> tab.", target: vsTab("alarms"), do: "click", until: vsTabShown("#vs_ah_high") },
            { title: "Four alarm levels", msg: "<b>High</b> and <b>Low</b> are critical levels. <b>Pre-High</b> and <b>Pre-Low</b> are early warnings. Each is a percent full.", target: "#vs_ah_high", do: "read" },
            { title: "Turn on Pre-High", msg: "Check <b>Pre-High Alarm Enabled</b>.", target: "#vs_ah_prehigh", do: "check" },
            { title: "Pre-High level", msg: "Type <b>75</b> percent.", target: "#vs_ah_prehigh_pct", do: "type", value: "75" },
            { title: "Turn on Pre-Low", msg: "Check <b>Pre-Low Alarm Enabled</b>.", target: "#vs_ah_prelow", do: "check" },
            { title: "Pre-Low level", msg: "Type <b>25</b> percent.", target: "#vs_ah_prelow_pct", do: "type", value: "25" }
            ],
            saveAndCloseVessel
        ]) });

    L.push({ id: "email-setup", section: "alerts", title: "Email Setup", manual: "5.4.3",
        desc: "Point Binventory at your SMTP mail server.",
        outro: "Binventory can now send email. Your IT department or internet provider supplies the server name and any login.",
        steps: flat([
            menu("Specifications", '[data-action="email-setup"]', "Email Setup", "Email Setup"),
            [
            { title: "Turn on email", msg: "Check <b>Enable SMTP Email Client</b>.", target: "#esEnable", do: "check" },
            { title: "Mail server", msg: "Type your SMTP server name. For practice type <b>mail.example.com</b>.", target: "#esServer", do: "type", value: "mail.example.com" },
            { title: "Port and security", msg: "Port 25 is the usual default. Your IT department will tell you if a secure connection or a different port is needed.", target: "#esPort", do: "read" },
            { title: "Authentication", msg: "If your server needs a login, check <b>Use SMTP Authentication</b> and fill in User ID and both password boxes.", target: "#esAuth", do: "read" },
            { title: "Email addresses", msg: "The manual recommends keeping the default Administrator and From addresses when you can.", target: "#esAdmin", do: "read" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#esSave", do: "click", until: function () { return !win("Email Setup"); } }
            ]
        ]) });

    L.push({ id: "contacts", section: "alerts", title: "Adding an email contact", manual: "5.4.4",
        desc: "Add a person who can receive alarm and report emails.",
        outro: "Contacts only get email after you assign them to a vessel's notifications or to a Site Status Report.",
        steps: flat([
            menu("Specifications", '[data-action="contact-maintenance"]', "Contact Maintenance", "Contact Maintenance"),
            [
            { title: "Add a contact", msg: "Click <b>Add New</b>.", target: "#cmAddNew", do: "click", until: function () { return visible("#csEmail"); } },
            { title: "First name", msg: "Type <b>Pat</b>.", target: "#csFirstName", do: "type", value: "Pat" },
            { title: "Last name", msg: "Type <b>Miller</b>.", target: "#csLastName", do: "type", value: "Miller" },
            { title: "Email address", msg: "The email address is required. Type <b>pat.miller@example.com</b>.", target: "#csEmail", do: "type", value: "pat.miller@example.com" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#csSave", do: "click", until: function () { return !visible("#csEmail"); } },
            { title: "Close Contact Maintenance", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("Contact Maintenance"); } }
            ]
        ]) });

    L.push({ id: "email-alerts", section: "alerts", title: "Email alerts for a vessel", manual: "5.4.7.6",
        desc: "Choose who gets emailed when a vessel alarms.",
        outro: "Jane Operator will now get an email when Cem Silo 1 reaches a high or low alarm level.",
        steps: flat([
            openVessel("v1", "Cem Silo 1"),
            [
            { title: "Email Notifications tab", msg: "Click the <b>Email</b> tab.", target: vsTab("email"), do: "click", until: vsTabShown("#vs_email_en") },
            { title: "Turn on notifications", msg: "Check <b>Enable Email Notifications</b>.", target: "#vs_email_en", do: "check" },
            { title: "Pick contacts", msg: "Click <b>Select Contacts</b>.", target: "#vs_sel_contacts", do: "click", until: function () { return shown("backdropAssignContacts"); } },
            { title: "Choose Jane", msg: "Click <b>Jane Operator</b> in Unassigned Contacts.", target: '#ac_tbody_un tr[data-cid="c1"]', do: "click" },
            { title: "Add her", msg: "Click <b>Add</b>.", target: "#ac_btn_add", do: "click", until: function () { return !!document.querySelector('#ac_tbody_as tr[data-cid="c1"]'); } },
            { title: "Close the contact list", msg: "Click <b>Close</b>.", target: "#ac_btn_close", do: "click", until: function () { return !shown("backdropAssignContacts"); } },
            { title: "Select Jane", msg: "Click <b>Jane</b> in the contact list. Each contact can get different emails.", target: '#vs_email_contacts_tbody tr[data-cid="c1"]', do: "click" },
            { title: "High alarm email", msg: "Check <b>High Alarm Enabled</b>.", target: "#vs_ef_high", do: "check" },
            { title: "Low alarm email", msg: "Check <b>Low Alarm Enabled</b>.", target: "#vs_ef_low", do: "check" }
            ],
            saveAndCloseVessel
        ]) });

    /* ======================= READINGS ======================= */
    L.push({ id: "measure", section: "setup", title: "Confirm a reading", manual: "5.2",
        desc: "Measure one vessel and follow its status.",
        outro: "A SmartBob reading goes Descending, then Retracting, and the panel updates with the new level.",
        steps: [
            { title: "Measure a vessel", msg: "Click <b>Measure</b> on Cem Silo 1.", target: vesselTile("v1") + ' [data-vessel-action="measure"]', do: "click" },
            { title: "Watch the status", msg: "The SmartBob is <b>Descending</b>: it is dropping its bob to find the product. Wait for it to start back up.",
              target: vesselTile("v1") + " .vessel-status-val", do: "wait",
              until: function () { return /retract/i.test(textOf(vesselTile("v1") + " .vessel-status-val")); } },
            { title: "Coming back up", msg: "<b>Retracting</b> means the bob is being raised. Wait for the reading to finish.",
              target: vesselTile("v1") + " .vessel-status-val", do: "wait",
              until: function () { var s = textOf(vesselTile("v1") + " .vessel-status-val"); return /retracted|ready/i.test(s); } },
            { title: "New reading", msg: "The height, volume and weight now show the new reading, with its date and time at the bottom of the panel.", target: vesselTile("v1"), do: "read" },
            { title: "Measure every vessel", msg: "You can also measure all vessels at once from the Vessel menu. Click <b>Vessel</b>.",
              target: function () { return menuBtn("Vessel"); }, do: "click", opensMenu: true, until: function () { return menuOpen("Vessel"); } },
            { title: "Measure All Vessels", msg: "This is <b>Measure All Vessels</b>. Click it to start a reading on every vessel.", target: "#menuItemMeasureAll", do: "click" }
        ] });

    L.push({ id: "vessel-details", section: "readings", title: "Viewing vessel details", manual: "5.3",
        desc: "Open the details screen for a vessel and its sensor.",
        outro: "Vessel Details shows contents, level, alarm settings, next measurement time and sensor information.",
        steps: [
            { title: "Open details", msg: "Click the level graphic of <b>Cem Silo 1</b>.", target: vesselTile("v1") + " .silo", do: "click", until: function () { return win("Vessel Details"); } },
            { title: "Vessel tab", msg: "The Vessel tab shows the vessel, its contents, product level, alarm settings and the time of the next measurement.", target: "#backdropApp .modal-cap", do: "read", ring: true },
            { title: "Sensors tab", msg: "Click the <b>Sensors</b> tab.", target: '#backdropApp [data-vd-tab="sensors"]', do: "click" },
            { title: "Sensor information", msg: "The sensor network and sensor type are at the top. This page helps when checking a sensor's setup with BinMaster technical support.", target: "#backdropApp .modal-cap", do: "read" },
            { title: "Close", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("Vessel Details"); } }
        ] });

    L.push({ id: "groups", section: "readings", title: "Creating a vessel group", manual: "5.4.8",
        desc: "Group vessels so they can be measured on a schedule.",
        outro: "Groups exist to organize vessels for measurement schedules. Next, attach a group to a schedule.",
        steps: flat([
            menu("Vessel", '[data-action="group-maintenance"]', "Vessel Group Maintenance", "Group Maintenance"),
            [
            { title: "Add a group", msg: "Click <b>Add New</b>.", target: "#gmBtnAdd", do: "click", until: function () { return visible("#gsName"); } },
            { title: "Group name", msg: "Use a unique, descriptive name. Type <b>Cement Silos</b>.", target: "#gsName", do: "type", value: "Cement Silos" },
            { title: "Assign vessels", msg: "Click <b>Assign Vessels</b>.", target: "#gsAssign", do: "click", until: function () { return visible("#vgaAdd"); } },
            { title: "Pick a vessel", msg: "Click <b>Cem Silo 1</b> in Unassigned Vessels.", target: '#vgaUnBody tr[data-vessel-id="v1"]', do: "click" },
            { title: "Add it", msg: "Click <b>Add</b>.", target: "#vgaAdd", do: "click", until: function () { return !!document.querySelector('#vgaAsBody tr[data-vessel-id="v1"]'); } },
            { title: "Pick another", msg: "Click <b>Cem Silo 2</b>.", target: '#vgaUnBody tr[data-vessel-id="v2"]', do: "click" },
            { title: "Add it", msg: "Click <b>Add</b>.", target: "#vgaAdd", do: "click", until: function () { return !!document.querySelector('#vgaAsBody tr[data-vessel-id="v2"]'); } },
            { title: "Done assigning", msg: "Click <b>Close</b>.", target: "#vgaClose", do: "click", until: function () { return !visible("#vgaAdd"); } },
            { title: "Save the group", msg: "Click <b>Save</b>.", target: "#gsSave", do: "click", until: function () { return !visible("#gsName"); } },
            { title: "Close", msg: "Click <b>Close</b>.", target: "#gmBtnClose", do: "click", until: function () { return !win("Group Maintenance"); } }
            ]
        ]) });

    L.push({ id: "schedules", section: "readings", title: "Creating a measurement schedule", manual: "5.4.9",
        desc: "Measure a group every hour on weekdays.",
        outro: "The schedule is active as soon as it is saved. Uncheck Schedule Active later to pause it.",
        steps: flat([
            menu("Vessel", '[data-action="schedule-maintenance"]', "Schedule Maintenance", "Scheduler Maintenance"),
            [
            { title: "Add a schedule", msg: "Click <b>Add New</b>.", target: "#backdropApp #smBtnAdd", do: "click", until: function () { return visible("#mssName"); } },
            { title: "Schedule name", msg: "Type <b>Weekday hourly</b>.", target: "#mssName", do: "type", value: "Weekday hourly" },
            { title: "Active by default", msg: "A new schedule is always active. You can pause it later by editing it.", target: "#mssActive", do: "read" },
            { title: "Regular intervals", msg: "Choose <b>Regular Intervals</b>.", target: '#backdropAppStack input[type="radio"][value="1"]', do: "check" },
            { title: "Skip Sunday", msg: "Uncheck <b>Sun</b> so it only runs on weekdays.", target: '#backdropAppStack input[data-i="0"]', do: "uncheck" },
            { title: "Skip Saturday", msg: "Uncheck <b>Sat</b>.", target: '#backdropAppStack input[data-i="6"]', do: "uncheck" },
            { title: "Every hour", msg: "Intervals is in minutes. <b>60</b> means every hour, from the Start Time to the End Time: 8 AM to 5 PM.", target: "#mssIntervalMinutes", do: "read" },
            { title: "Assign groups", msg: "Click <b>Assign Groups</b>.", target: "#mssAssign", do: "click", until: function () { return visible("#sgaAdd"); } },
            { title: "Pick a group", msg: "Click <b>Mill Line A</b>.", target: '#sgaUnBody tr[data-group-id="g1"]', do: "click" },
            { title: "Add it", msg: "Click <b>Add</b>.", target: "#sgaAdd", do: "click", until: function () { return !!document.querySelector('#sgaAsBody tr[data-group-id="g1"]'); } },
            { title: "Done assigning", msg: "Click <b>Close</b>.", target: "#sgaClose", do: "click", until: function () { return visible("#mssSave"); } },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#mssSave", do: "click", until: function () { return !visible("#mssName"); } },
            { title: "Close", msg: "Click <b>Close</b>.", target: "#backdropApp #smBtnClose", do: "click", until: function () { return !win("Scheduler Maintenance"); } }
            ]
        ]) });

    L.push({ id: "temporary-group", section: "readings", title: "Showing only some vessels", manual: "5.4.12",
        desc: "Use Temporary Group Setup to hide vessels for now.",
        outro: "Temporary groups last until you log out or exit. Use >> to show every vessel again.",
        steps: flat([
            menu("Vessel", '[data-action="temporary-group"]', "Temporary Group Setup", "Temporary Group"),
            [
            { title: "Hide everything", msg: "The right list is what shows on the main screen. Click <b>&lt;&lt;</b> to move every vessel to the left.", target: "#tgBtnAllRem", do: "click", until: function () { var s = el("#tgSel"); return !!s && s.options.length === 0; } },
            { title: "Pick a vessel", msg: "In Available Vessels, select <b>Cem Silo 1</b>.", target: "#tgAvail", do: "select", value: "Cem Silo 1" },
            { title: "Show it", msg: "Click <b>&gt;</b> to move it to Selected Vessels.", target: "#tgBtnSel", do: "click", until: function () { var s = el("#tgSel"); return !!s && s.options.length >= 1; } },
            { title: "Pick another", msg: "Select <b>Cem Silo 2</b>.", target: "#tgAvail", do: "select", value: "Cem Silo 2" },
            { title: "Show it", msg: "Click <b>&gt;</b>.", target: "#tgBtnSel", do: "click", until: function () { var s = el("#tgSel"); return !!s && s.options.length >= 2; } },
            { title: "Accept", msg: "Click <b>OK</b>.", target: "#tgBtnOk", do: "click", until: function () { return !win("Temporary Group"); } },
            { title: "Only two vessels", msg: "The main screen now shows just the vessels you picked.", target: "#vesselGrid", do: "read" }
            ]
        ]) });

    /* ======================= REPORTS ======================= */
    function printReports(label) {
        return flat([
            menu("Reports", '[data-modal="print-reports"]', "Print Reports", "Print Reports"),
            [
            { title: "Pick the report", msg: "Choose <b>" + label + "</b>.", target: "#reportSelect", do: "select", value: label },
            { title: "Open it", msg: "Click <b>Select</b>.", target: "#btnPrintSelect", do: "click", until: function () { return win(label) || win("Inventory Report") || !shown("backdropPrint"); } }
            ]
        ]);
    }
    L.push({ id: "inventory-report", section: "reports", title: "Inventory Report", manual: "5.4.13.2",
        desc: "A table of every vessel's level.",
        outro: "In the full product the report opens in Crystal Reports, where it can be printed or saved as a PDF.",
        steps: flat([
            printReports("Inventory Report"),
            [
            { title: "Output type", msg: "<b>Report</b> makes a table. Bar Graph and CSV export are the other choices.", target: "#invOutReport", do: "check" },
            { title: "Include every vessel", msg: "Vessels in the right list are included. Click <b>&gt;&gt;</b> to include all of them.", target: "#invBtnSelAll", do: "click", until: function () { var s = el("#invSelList"); return !!s && s.options.length > 1; } },
            { title: "Sort order", msg: "The report can be sorted by vessel name or by contents.", target: "#invSortOrder", do: "read" },
            { title: "Create it", msg: "Click <b>Create</b>.", target: "#invCreate", do: "click", until: function () { return win("Inventory Report") && shown("backdropApp"); } },
            { title: "Your report", msg: "Each vessel is listed with percent full, volume, weight and status.", target: "#backdropApp .modal-cap", do: "read" },
            { title: "Close", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !shown("backdropApp"); } }
            ]
        ]) });

    L.push({ id: "history-report", section: "reports", title: "Inventory History Report", manual: "5.4.13.3",
        desc: "Chart one vessel's level over its recent readings.",
        outro: "History charts can use the last 30 or 60 readings or a date range, for product or headroom.",
        steps: flat([
            printReports("Inventory History Report"),
            [
            { title: "Range of data", msg: "Choose <b>Last 30 Measurements</b>.", target: "#histRad30", do: "check" },
            { title: "Vessel", msg: "Choose <b>Cem Silo 2</b>.", target: "#histVesselList", do: "select", value: "Cem Silo 2" },
            { title: "Vertical scale", msg: "Choose <b>Product Weight</b>.", target: "#histChartYAxis", do: "select", value: "Product Weight" },
            { title: "Create the chart", msg: "Click <b>Create</b>.", target: "#histCreate", do: "click" },
            { title: "Your chart", msg: "This is the level history for Cem Silo 2.", target: topCaption, do: "read" }
            ]
        ]) });

    L.push({ id: "site-status-report", section: "reports", title: "Emailed Site Status Report", manual: "5.4.14",
        desc: "Schedule an email of the latest readings for every vessel at a site.",
        outro: "Site Status Reports are scheduled like measurements, but you assign contacts to email instead of vessels to measure.",
        steps: flat([
            menu("Reports", '[data-action="email-reports"]', "Email Reports", "Email Reports"),
            [
            { title: "Pick the report", msg: "Click <b>Site Status Report</b>.", target: "#erOptSite", do: "click" },
            { title: "Open it", msg: "Click <b>Select</b>.", target: "#erBtnSelect", do: "click", until: function () { return win("Site Status Reports"); } },
            { title: "Add a schedule", msg: "Click <b>Add New</b>.", target: "#backdropAppStack #smBtnAdd", do: "click", until: function () { return visible("#mssName"); } },
            { title: "Name it", msg: "Type <b>Morning report</b>.", target: "#mssName", do: "type", value: "Morning report" },
            { title: "Daily", msg: "Choose <b>Daily</b>.", target: '#backdropAppStack2 input[type="radio"][value="2"]', do: "check" },
            { title: "Run on Mon", msg: "Daily reports go out on the days you check. Check <b>Mon</b>.", target: '#backdropAppStack2 input.mss-day[data-i="1"]', do: "check" },
            { title: "Run on Tue", msg: "Check <b>Tue</b>.", target: '#backdropAppStack2 input.mss-day[data-i="2"]', do: "check" },
            { title: "Run on Wed", msg: "Check <b>Wed</b>.", target: '#backdropAppStack2 input.mss-day[data-i="3"]', do: "check" },
            { title: "Run on Thu", msg: "Check <b>Thu</b>.", target: '#backdropAppStack2 input.mss-day[data-i="4"]', do: "check" },
            { title: "Run on Fri", msg: "Check <b>Fri</b>.", target: '#backdropAppStack2 input.mss-day[data-i="5"]', do: "check" },
            { title: "Start time", msg: "The report goes out at the Start Time shown here. Use the arrows to set it, for example 8 AM.", target: "#mssInfoStartTimeDisplay", do: "read" },
            { title: "Who gets it", msg: "In Binventory this button assigns the <b>contacts</b> who receive the report. In this practice copy it opens the measurement group list instead, so leave it for now.", target: "#mssAssign", do: "read" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#mssSave", do: "click", until: function () { return !visible("#mssName"); } },
            { title: "Close", msg: "Click <b>Close</b>.", target: "#backdropAppStack #smBtnClose", do: "click", until: function () { return !win("Site Status Reports"); } }
            ]
        ]) });

    /* ======================= USERS AND BACKUP ======================= */
    L.push({ id: "users", section: "admin", title: "Adding a user", manual: "5.4.10",
        desc: "Create a login and pick the right access level.",
        outro: "Users only need to log in once System Setup no longer has Do NOT require users to login checked.",
        steps: flat([
            menu("File", '[data-action="user-maintenance"]', "User Maintenance", "User Maintenance"),
            [
            { title: "Add a user", msg: "Click <b>Add New</b>.", target: "#umAddNew", do: "click", until: function () { return visible("#usUserId"); } },
            { title: "First name", msg: "Type <b>Sam</b>.", target: "#usFirstName", do: "type", value: "Sam" },
            { title: "Last name", msg: "Type <b>Lee</b>.", target: "#usLastName", do: "type", value: "Lee" },
            { title: "User ID", msg: "User IDs must be unique. Type <b>slee</b>.", target: "#usUserId", do: "type", value: "slee" },
            { title: "Access level", msg: "<b>System User (Read-Only)</b> can view levels, measure and run reports. <b>System User</b> can also change settings under the Vessel menu. <b>Administrator User</b> has all access. Choose <b>System User</b>.", target: "#usSecurity", do: "select", value: "System User" },
            { title: "Password", msg: "Type <b>4321</b>.", target: "#usPassword", do: "type", value: "4321" },
            { title: "Verify password", msg: "Type <b>4321</b> again.", target: "#usVerifyPassword", do: "type", value: "4321" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#usSave", do: "click", until: function () { return !visible("#usUserId"); } },
            { title: "Close", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("User Maintenance"); } }
            ]
        ]) });

    L.push({ id: "export", section: "admin", title: "Backing up with System Export", manual: "5.4.15",
        desc: "Save your whole configuration to a DAT file.",
        outro: "Export again after any configuration change, such as adding a vessel. Only administrators can export or import.",
        steps: flat([
            menu("File", '[data-action="export-data"]', "System Export", "System Export"),
            [
            { title: "File name", msg: "The default file name uses today's date and the DAT extension. Use <b>Browse</b> if you want a different place.", target: "#simSysTxtFileLocation", do: "read" },
            { title: "Export", msg: "Click <b>Export</b>.", target: "#simSysBtnPrimary", do: "click", until: function () { return shown("backdropBinventoryMsg"); } },
            { title: "Done", msg: "The export is complete. Click <b>OK</b>.", target: "#binvMsgBtnOk", do: "click", until: function () { return !shown("backdropBinventoryMsg"); } }
            ]
        ]) });

    /* ======================= TROUBLESHOOTING (built-in scenarios) ======================= */
    L.push({ id: "ts-db-read-only", section: "trouble", title: "Status: Database Read Only", manual: "5.2, 9.2",
        desc: "Vessels show but cannot be measured. Find the server IP address and set it as the Host IP Address in Site Setup.",
        scenario: "db-read-only", stepCount: 20, minutes: 6 });
    L.push({ id: "ts-pending-unknown", section: "trouble", title: "Status: Pending or Unknown", manual: "5.2",
        desc: "Readings never finish. Restart the Engine and Scheduler services, then measure again.",
        scenario: "pending-unknown-status", stepCount: 23, minutes: 5 });


    /* ======================= SETUP: before and during installation (reading lessons, doc.html) ======================= */
    L.push({ id: "plan", section: "setup", kind: "doc", title: "Before you install", manual: "2, 3, 4.1, 4.2, 4.4.1",
        desc: "What Binventory is made of, what the PC needs and which setup to choose.",
        outro: "You know the parts of Binventory, what the PC needs and whether to use Easy or Advanced setup.",
        doc: [
            { title: "Binventory has three parts", html: "<p>A full installation has three parts. They can all be on one PC or spread across a local network.</p><ol><li><b>Workstation.</b> The window you use to see vessels, set things up and run reports. There can be more than one.</li><li><b>Engine and Scheduler services.</b> Background programs that start with the PC. They talk to the sensors, run measurement schedules and send email alerts. Each site of vessels needs its own running set.</li><li><b>SQL Server database.</b> Stores all settings and readings, using Microsoft SQL Server 2014 Express. One system has one database.</li></ol>" },
            { title: "Stand-alone or networked", html: "<p><b>Single stand-alone.</b> Everything on one PC. The easiest to install, but only usable on that PC.</p><p><b>Networked.</b> Parts on several PCs on a private network. More work to set up, but it supports any number of sites and remote Workstations, for example a database at the office and a site at each plant.</p>" },
            { title: "Check the PC", html: "<ul><li>Windows 11, 10, 8.1, 8 or 7 SP1, or Windows Server 2008 R2 through 2022</li><li>Microsoft .NET Framework 3.5 SP1 and 4.0. The installer finds any that are missing.</li><li>A 1.4 GHz 64-bit or 1.0 GHz 32-bit processor. 2.0 GHz or faster is recommended.</li><li>4 GB of RAM and 7 GB of free disk space</li><li>An 800 by 600 screen. 1024 by 768 or better is recommended.</li><li>A free USB port for the installation flash drive</li></ul><p>Binventory cannot be guaranteed to work on a PC that does not meet these.</p>" },
            { title: "A way to reach the sensors", html: "<p>The PC that hosts a site needs at least one connection to the sensor network:</p><ul><li>an RS-232 serial port plus an RS-232 to RS-485 interface,</li><li>a USB port plus a USB to serial adaptor, or</li><li>a network connection plus a serial device server.</li></ul><p>Email alerts also need network or internet access to an SMTP email server. Networked systems need every PC on the same IP network.</p>" },
            { title: "Easy or Advanced setup", html: "<p>The installer asks for a setup type. <b>Easy</b> installs every part. <b>Advanced</b> lets you pick parts.</p><p>Use <b>Easy</b> when:</p><ul><li>you are installing a single stand-alone system, or</li><li>this PC in a networked system will host both the database and a set of services.</li></ul><p>Every other install in a networked system uses <b>Advanced</b>.</p>" },
            { title: "Networked systems: make a list first", html: "<p>Before a network install, list every PC you will use. For each one write down:</p><ul><li>its computer name or static IP address</li><li>its installation type: Database, Database with Site, Site, or Remote Workstation</li><li>the COM port, if it hosts a site of vessels</li><li>the Windows administrator user name and password</li></ul><p>A PC that hosts a site needs a free serial port and should be fairly close to the vessels.</p>" }
        ] });

    L.push({ id: "install", section: "setup", kind: "doc", title: "Installing on one PC", manual: "4.3",
        desc: "The Easy setup for a single stand-alone system, up to first launch.",
        outro: "Binventory is installed. The next lessons set up the Workstation in the order the manual uses.",
        doc: [
            { title: "Log in as an administrator", html: "<p>Log in to Windows with an <b>administrator</b> account. Installation fails without administrator rights.</p><p>Make sure the PC meets the requirements from the previous lesson.</p>" },
            { title: "Run Binventory Setup", html: "<p>Insert the USB flash drive with the Binventory software. Find the <b>Binventory Setup</b> file on it and run it.</p>" },
            { title: "Choose Easy", html: "<p>Follow the on-screen instructions. When asked for the setup type, choose <b>Easy</b>. This installs the Workstation, the Engine and Scheduler services and the SQL database.</p><p>If the PC is missing a needed .NET Framework, the installer tells you.</p>" },
            { title: "Open the Workstation", html: "<p>When the installation finishes, double-click the <b>Binventory Workstation</b> icon on the desktop. It is also under <b>eBob</b> in the Windows start menu.</p><p>The first time, the main screen is mostly blank because no vessels are set up yet. A stand-alone system does not need Binventory Administrator because its defaults are already set.</p>" },
            { title: "What comes next", html: "<p>The manual sets up a new system in this order, and so do the next lessons:</p><ol><li>Tour the Workstation</li><li>System Setup and Site Setup</li><li>Add a sensor network for your sensor type</li><li>Add each vessel, then copy it for similar vessels</li><li>Take a test measurement</li></ol>" }
        ] });

    L.push({ id: "network-install", section: "setup", kind: "doc", title: "Installing a networked system", manual: "4.4, 6",
        desc: "Install order, which parts go on each PC and connecting each PC to the database.",
        outro: "Each PC knows where the database is, and each site's Engine service is assigned to its site.",
        doc: [
            { title: "Four kinds of install", html: "<ol><li><b>Database.</b> The SQL database and the Workstation, usually on a server.</li><li><b>Database with Site.</b> Everything on one PC.</li><li><b>Site.</b> The Engine and Scheduler services plus the Workstation. One for each site of vessels.</li><li><b>Remote Workstation.</b> Only the Workstation. Any number of these.</li></ol><p>A system has only one database.</p>" },
            { title: "Install the database first", html: "<p>Start on the PC or server that will host the database. Use <b>Advanced</b> setup and do a Database or Database with Site install.</p><p>Log in as an administrator, run Binventory Setup from the flash drive, choose Advanced and leave the destination folder at its default.</p>" },
            { title: "Pick the parts for each install", html: "<p>Advanced setup asks which features to install. <b>Application</b> is always checked.</p><ul><li><b>Database:</b> Application and SQL Database</li><li><b>Database with Site:</b> every feature</li><li><b>Site:</b> Application plus Engine and Scheduler Services</li><li><b>Remote Workstation:</b> Application only</li></ul><p>Wait for it to finish and click <b>Finish</b>.</p>" },
            { title: "Set up the sites", html: "<p>On the database PC, open the Workstation and add every site in <b>Site Maintenance</b>. You can also add each site's vessels now, or later from the site PC.</p><p>See the Site Setup lesson in this section.</p>" },
            { title: "Install each site PC", html: "<p>On each PC that hosts a site, do a <b>Site</b> install with Advanced setup.</p><p>Then open <b>Binventory Administrator</b> under eBob in the start menu:</p><ol><li>Set <b>Server Name</b> to the database PC's static IP address or computer name followed by \\EBOB, for example 192.168.254.100\\EBOB. Leave it as (local)\\EBOB only on the database PC.</li><li>Type the database password in <b>Verify Password</b>. The default is in manual section 6.</li><li>Click <b>Verify</b> and wait for the success message.</li><li>Set the permanent site assignment for the Engine service. On a PC with no site, choose Unassigned or engine service not installed.</li><li>Click <b>Save</b>, then reboot the PC.</li></ol>" },
            { title: "Point Workstations at the Engine", html: "<p>In <b>Site Setup</b>, each site's <b>Host IP Address</b> must be the static IP address or computer name of the PC running that site's Engine service. The default 127.0.0.1 only works on that PC itself.</p><p>A wrong Host IP Address is a common cause of the <b>Database Read Only</b> status on remote Workstations.</p>" },
            { title: "SQL Server and firewall", html: "<p>The installer now sets these up for you. For reference, a networked system needs:</p><ul><li>Named Pipes and TCP/IP enabled for the EBOB SQL Server instance</li><li>an inbound firewall rule for SQL Server (EBOB), the sqlservr.exe program</li><li>an inbound rule for SQL Server Browser on UDP port 1434</li><li>an inbound rule for the Binventory Engine Service on TCP port 8093</li></ul>" }
        ] });

    L.push({ id: "upgrade", section: "admin", kind: "doc", title: "Upgrading Binventory", manual: "7, 8, 9",
        desc: "Moving from IMS or eBob 4, or updating an earlier Binventory 5.",
        outro: "Always export your system before an upgrade so it can be restored.",
        doc: [
            { title: "From IMS or eBob 4", html: "<p>Binventory 5 does not install over IMS v2.2.3 or eBob 4. Leave the old program installed until Binventory 5 is fully set up, and check the PC meets the new requirements.</p><p>If every sensor is a SmartBob-II or SmartBob-TS1, Protocol A is recommended. SmartBob-II firmware must be 2.3 or newer, and 2.8 or newer for Max Drop.</p>" },
            { title: "Back up first", html: "<p>Binventory 5.4 installs over any earlier Binventory 5. Before you start, open the Workstation and use <b>System Export</b> under the File menu to save a backup with the default file name. Then close the Workstation.</p>" },
            { title: "Stand-alone upgrade", html: "<ol><li>Log in to Windows as an administrator.</li><li>Run Binventory Setup from the flash drive and follow the instructions.</li><li>Open Binventory Administrator, confirm the database settings, enter the database password and click <b>Verify</b>.</li><li>Set both site assignments to your existing site and click <b>Save</b>.</li><li>Reboot the PC.</li></ol>" },
            { title: "Networked upgrade", html: "<ol><li>Export the system, then note the Server Name in Binventory Administrator on a PC that does not host the database.</li><li>Collect the static IP address or computer name of each PC that hosts a site.</li><li>Close the Workstation on every PC.</li><li>Upgrade the database PC first, then verify the connection in Binventory Administrator, set its Engine site assignment, save and reboot.</li><li>On the database PC, open each site in Site Maintenance and set its <b>Host IP Address</b> to that site PC's address.</li><li>Upgrade every other PC the same way and reboot each one.</li></ol>" }
        ] });

    /* ======================= SETUP: copy a vessel ======================= */
    L.push({ id: "more-vessels", section: "setup", title: "Adding similar vessels quickly", manual: "4.3, 5.4.6",
        desc: "Copy a vessel with Add New Using, then sort the vessel list.",
        outro: "Add New Using copies every setting. Always change at least the name and the sensor address.",
        steps: flat([
            menu("Vessel", '[data-action="vessel-maintenance"]', "Vessel Maintenance", "Vessel Maintenance"),
            [
            { title: "Pick a vessel to copy", msg: "Click <b>Cem Silo 1</b>.", target: '#backdropApp tr[data-vid="v1"]', do: "click" },
            { title: "Copy it", msg: "Click <b>Add New Using</b>. Vessel Setup opens with a copy of Cem Silo 1.", target: "#vmBtnAddUsing", do: "click", until: function () { return visible("#vs_name"); } },
            { title: "New name", msg: "At minimum, change the name and the sensor address. Type <b>Cem Silo 3</b>.", target: "#vs_name", do: "type", value: "Cem Silo 3" },
            { title: "Sensor tab", msg: "Click the <b>Sensor</b> tab.", target: vsTab("sensor"), do: "click", until: vsTabShown("#vs_sb_addr1") },
            { title: "New sensor address", msg: "The copy still has Cem Silo 1's address. Every sensor needs its own. Type <b>33</b>.", target: "#vs_sb_addr1", do: "type", value: "33" },
            { title: "Save", msg: "Click <b>Save</b>.", target: "#vsSave", do: "click", until: function () { return !visible("#vs_name"); } },
            { title: "Sort the list", msg: "Vessels show on the main screen in this list's order. Click <b>Sort by Name</b>.", target: "#vmBtnSortName", do: "click" },
            { title: "Fine tune the order", msg: "<b>Sort by Contents</b>, <b>Move Up</b> and <b>Move Down</b> also change the order. Move Up and Move Down work on the highlighted vessel.", target: "#vmBtnMoveUp", do: "read" },
            { title: "Close Vessel Maintenance", msg: "Click <b>Close</b>. The main screen updates.", target: "#vmBtnClose", do: "click", until: function () { return !win("Vessel Maintenance"); } }
            ]
        ]) });

    /* ======================= TROUBLESHOOTING (guided checks) ======================= */
    function btnIn(containerSel, text) {
        return function () {
            var c = document.querySelector(containerSel);
            while (c && !(c.classList && c.classList.contains("backdrop"))) c = c.parentElement;
            if (!c) return null;
            var bs = c.querySelectorAll("button");
            for (var i = bs.length - 1; i >= 0; i--) if (bs[i].textContent.trim() === text && visible(bs[i])) return bs[i];
            return null;
        };
    }
    var WIRING = '<span class="eg-links">' +
        '<a href="https://support.binmaster.com/kb/view/3ae8f673-3dfd-466f-b9fd-2e814996426a" target="_blank" rel="noopener">SmartBob wiring</a>' +
        '<a href="https://support.binmaster.com/kb/view/9ab4c1b8-9795-4e4a-90a7-192da74a5729" target="_blank" rel="noopener">CNCR wiring</a>' +
        '<a href="https://support.binmaster.com/kb/view/a5d3b4e0-a336-43f4-94f4-962eeef7d7af" target="_blank" rel="noopener">NCR wiring</a></span>';
    var vesselSetupCancel = [
        { title: "Leave without changes", msg: "Click <b>Cancel</b>.", target: "#vsCancel", do: "click", until: function () { return !visible("#vs_name"); } },
        { title: "Close Vessel Maintenance", msg: "Click <b>Close</b>.", target: "#vmBtnClose", do: "click", until: function () { return !win("Vessel Maintenance"); } }
    ];

    L.push({ id: "ts-comm-error", section: "trouble", title: "Status: Communications Error", manual: "5.2, 5.4.5, 5.4.7.4",
        desc: "The sensor did not answer. Check the wiring, the network settings and the sensor address.",
        outro: "If the wiring, the network settings and the address all check out and the error stays, call BinMaster Technical Support at 402-434-9105.",
        steps: flat([
            [
            { title: "What the status means", msg: "<b>Communications Error</b> means Binventory did not get the right reply from the sensor on its last measurement. It is usually RS-485 wiring, but it can be software setup or hardware.", target: vesselTile("v1") + " .vessel-status-val", do: "read" },
            { title: "Check the wiring first", msg: "Sensors on a network are daisy-chained on shielded twisted-pair cable, with matching polarity, both ends terminated and only one bias point. Open the wiring guide for your sensor:" + WIRING, target: vesselTile("v1"), do: "read" }
            ],
            menu("Specifications", '[data-action="sensor-networks"]', "Sensor Networks", "Sensor Networks"),
            [
            { title: "Pick the network", msg: "Click <b>COM2 Sensor Network</b>.", target: '#backdropApp tr[data-nid="n1"]', do: "click" },
            { title: "Open its setup", msg: "Click <b>Select</b>.", target: "#snBtnSelect", do: "click", until: function () { return shown("backdropSnSetup"); } },
            { title: "Protocol", msg: "Every sensor on the network must use this protocol. SmartBob-II uses <b>Protocol A</b>. Original SmartBobs use Protocol B. Most other BinMaster sensors use Modbus/RTU.", target: "#snsProtocol", do: "read" },
            { title: "Port", msg: "This must be the COM port the RS-485 interface is plugged into. For a serial device server, check the IP address and port instead.", target: "#snsCom", do: "read" },
            { title: "Communication parameters", msg: "Baud rate, data bits, parity and stop bits must match the sensors. Leave the advanced settings alone unless BinMaster support asks.", target: "#snsBaud", do: "read" },
            { title: "Leave without changes", msg: "Click <b>Cancel</b>.", target: "#snSetupCancel", do: "click", until: function () { return !shown("backdropSnSetup"); } },
            { title: "Close Sensor Networks", msg: "Click <b>Close</b>.", target: closeBtn("backdropApp"), do: "click", until: function () { return !win("Sensor Networks"); } }
            ],
            openVessel("v1", "Cem Silo 1"),
            [
            { title: "Sensor tab", msg: "Click the <b>Sensor</b> tab.", target: vsTab("sensor"), do: "click", until: vsTabShown("#vs_sb_addr1") },
            { title: "Right network", msg: "The vessel must be on the network its sensor is wired to.", target: "#vs_network", do: "read" },
            { title: "Enabled", msg: "<b>Enabled</b> must be checked, or the sensor is skipped.", target: "#vs_sb_en1", do: "read" },
            { title: "Sensor address", msg: "The address must match the address set on the sensor itself, and no two sensors on a network can share one. Valid addresses are 1 to 255.", target: "#vs_sb_addr1", do: "read" }
            ],
            vesselSetupCancel,
            [
            { title: "Test it", msg: "After fixing anything, click <b>Measure</b> to test the sensor.", target: vesselTile("v1") + ' [data-vessel-action="measure"]', do: "click" }
            ]
        ]) });

    L.push({ id: "ts-readings-wrong", section: "trouble", title: "Readings look wrong", manual: "5.4.7.2, 5.4.7.3, 5.4.7.4",
        desc: "Level, volume or weight do not match the vessel. Check its shape, size, density and offset.",
        outro: "Volume comes from the vessel type and dimensions, and weight comes from the volume and density. The sensor offset shifts every level reading.",
        steps: flat([
            [
            { title: "Start with the numbers", msg: "Cem Silo 1 shows its level, volume and weight here. If one looks wrong, check the vessel's setup.", target: vesselTile("v1"), do: "read" }
            ],
            openVessel("v1", "Cem Silo 1"),
            [
            { title: "Type/Shape tab", msg: "Click the <b>Type/Shape</b> tab.", target: vsTab("typeshape"), do: "click", until: vsTabShown("#vs_vessel_type") },
            { title: "Vessel type", msg: "The type must match the real vessel. A cone or a sloped bottom left out makes the volume wrong.", target: "#vs_vessel_type", do: "read" },
            { title: "Height", msg: "Check <b>Height</b> against the drawings or a tape measure. It is in the site's distance units, feet here.", target: "#vs_sp_0", do: "read" },
            { title: "Width", msg: "Check the <b>Width</b> too. A small error in width makes a big error in volume.", target: "#vs_sp_1", do: "read" },
            { title: "Contents tab", msg: "Click the <b>Contents</b> tab.", target: vsTab("contents"), do: "click", until: vsTabShown("#vs_density_val") },
            { title: "Density", msg: "Unless the vessel uses a custom lookup table, an accurate density is critical. Weight is volume times density. This one says 40 pounds per cubic foot, which is low for cement.", target: "#vs_density_val", do: "read" },
            { title: "Fix the density", msg: "Use the density from your supplier. For practice, type <b>94</b>.", target: "#vs_density_val", do: "type", value: "94" },
            { title: "Or specific gravity", msg: "You can enter <b>Specific Gravity</b> instead of density if that is what you have.", target: "#vs_den_sg", do: "read" },
            { title: "Sensor tab", msg: "Click the <b>Sensor</b> tab.", target: vsTab("sensor"), do: "click", until: vsTabShown("#vs_sb_offset1") },
            { title: "Sensor offset", msg: "For a top-down sensor, the offset is the distance from the sensor to the top of the vessel or the full mark. Positive is down inside the vessel. It defaults to zero.", target: "#vs_sb_offset1", do: "read" }
            ],
            saveAndCloseVessel,
            [
            { title: "Measure again", msg: "Binventory uses the new density from the next reading on. Click <b>Measure</b> on Cem Silo 1. In this practice copy the weight shown does not change.", target: vesselTile("v1") + ' [data-vessel-action="measure"]', do: "click" }
            ]
        ]) });

    L.push({ id: "ts-no-schedule", section: "trouble", title: "Scheduled readings are not happening", manual: "2, 5.4.8, 5.4.9",
        desc: "Check that the schedule is active, runs on the right days and has groups assigned.",
        outro: "A schedule only measures vessels in its assigned groups, on its checked days, and only while the site's Scheduler service is running.",
        steps: flat([
            menu("Vessel", '[data-action="schedule-maintenance"]', "Schedule Maintenance", "Scheduler Maintenance"),
            [
            { title: "Pick the schedule", msg: "Click <b>Morning sweep</b>.", target: '#backdropApp tr[data-schedule-id="s1"]', do: "click" },
            { title: "Open it", msg: "Click <b>Select</b>.", target: "#backdropApp #smBtnSelect", do: "click", until: function () { return visible("#mssName"); } },
            { title: "Schedule Active", msg: "<b>Schedule Active</b> must be checked. Unchecking it pauses the schedule.", target: "#mssActive", do: "read" },
            { title: "Days", msg: "The schedule only runs on the days checked under <b>Run only on</b>.", target: '#backdropAppStack input.mss-day[data-i="1"]', do: "read" },
            { title: "Times", msg: "With Regular Intervals it measures every this many minutes between the start and end times.", target: "#mssIntervalMinutes", do: "read" },
            { title: "Assigned groups", msg: "Schedules measure groups, not single vessels. Click <b>Assign Groups</b>.", target: "#mssAssign", do: "click", until: function () { return visible("#sgaClose"); } },
            { title: "Groups on this schedule", msg: "Only vessels in these groups are measured. A vessel in no assigned group is never measured by this schedule.", target: "#sgaAsBody", do: "read" },
            { title: "Close the group list", msg: "Click <b>Close</b>.", target: "#sgaClose", do: "click", until: function () { return visible("#mssCancel"); } },
            { title: "Leave without changes", msg: "Click <b>Cancel</b>.", target: "#mssCancel", do: "click", until: function () { return !visible("#mssName"); } },
            { title: "Close Schedule Maintenance", msg: "Click <b>Close</b>.", target: "#backdropApp #smBtnClose", do: "click", until: function () { return !win("Scheduler Maintenance"); } },
            { title: "The Scheduler service", msg: "Schedules are run by the site's <b>Scheduler service</b> on the PC hosting the site. If that PC is off or the service is stopped, nothing is measured. A vessel whose sensor is not Enabled is also skipped.", target: titleBar, do: "read" }
            ]
        ]) });

    L.push({ id: "ts-no-email", section: "trouble", title: "Email alerts are not arriving", manual: "5.4.3, 5.4.7.6",
        desc: "Check the SMTP server in Email Setup and the vessel's email settings.",
        outro: "Binventory needs an SMTP server to send email, and each vessel needs notifications enabled for the right contacts.",
        steps: flat([
            menu("Specifications", '[data-action="email-setup"]', "Email Setup", "Email Setup"),
            [
            { title: "Email turned on", msg: "Email must be enabled here.", target: "#esEnable", do: "read" },
            { title: "SMTP server", msg: "At minimum Binventory needs your network's SMTP server name. Here it is <b>blank</b>, so no email can be sent. Ask your IT department or internet provider for it.", target: "#esServer", do: "read" },
            { title: "Port and security", msg: "The port and security settings must match the server. Port 25 is the default.", target: "#esPort", do: "read" },
            { title: "Authentication", msg: "If the server requires a login, check <b>Use SMTP Authentication</b> and enter the User ID and password.", target: "#esAuth", do: "read" },
            { title: "Leave without changes", msg: "Click <b>Cancel</b>.", target: btnIn("#esSave", "Cancel"), do: "click", until: function () { return !win("Email Setup"); } }
            ],
            openVessel("v1", "Cem Silo 1"),
            [
            { title: "Email Notifications tab", msg: "Click the <b>Email</b> tab.", target: vsTab("email"), do: "click", until: vsTabShown("#vs_email_en") },
            { title: "Notifications enabled", msg: "<b>Enable Email Notifications</b> must be checked for this vessel.", target: "#vs_email_en", do: "read" },
            { title: "Who gets what", msg: "Each contact in this list needs the alarm types checked that they should get. If the list is empty, nobody is emailed.", target: "#vs_email_contacts_tbody", do: "read" }
            ],
            vesselSetupCancel
        ]) });

    /* final order inside each section */
    var ORDER = ["plan", "install", "network-install", "tour", "system-setup", "site-setup", "sensor-network", "add-vessel", "more-vessels", "measure",
        "level-alarms", "email-setup", "contacts", "email-alerts",
        "vessel-details", "groups", "schedules", "temporary-group",
        "inventory-report", "history-report", "site-status-report",
        "users", "export", "upgrade",
        "ts-comm-error", "ts-readings-wrong", "ts-no-schedule", "ts-no-email", "ts-pending-unknown", "ts-db-read-only"];
    L.sort(function (a, b) { return ORDER.indexOf(a.id) - ORDER.indexOf(b.id); });
    L.forEach(function (l) { if (l.id === "measure") { l.section = "setup"; } });

    L.forEach(function (l) {
        if (!l.stepCount) l.stepCount = l.doc ? l.doc.length : (l.steps ? l.steps.length : 0);
        if (!l.minutes) l.minutes = Math.max(1, Math.round(l.stepCount * 0.25));
    });

    window.EbobLessons = {
        sections: SECTIONS,
        lessons: L,
        byId: function (id) { for (var i = 0; i < L.length; i++) if (L[i].id === id) return L[i]; return null; },
        helpers: { el: el, visible: visible, win: win, shown: shown, menuBtn: menuBtn, menuOpen: menuOpen }
    };
})();
