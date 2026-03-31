(function () {
    "use strict";

    /** Legacy key; storage is cleared on load and not written (session-only state). */
    var STORAGE_KEY = "ebobSimState.v1";
    var GRID_SIZE = 16;
    /** Invalidates in-flight staggered vessel reveals when renderGrid runs again. */
    var vesselGridRevealGen = 0;

    /**
     * Simulated System Export/Import — fake Desktop\Documents (no real .dat on disk).
     * Documents list is empty until a successful export; then exactly one .dat appears for import.
     */
    var SIM_FAKE_DOCS_ROOT = "Desktop\\Documents";
    var simSysExportDatMeta = null;
    var simSysExportSnapshotJson = null;
    var simSysExportImportMode = 1;

    /** After simulated import + UAC, full reload restores state from sessionStorage (mirrors Application.Restart). */
    var SIM_IMPORT_RESTART_FLAG = "ebobSimImportRestartV1";
    var SIM_IMPORT_STATE_KEY = "ebobSimImportStateV1";
    var pendingImportRestartToast = false;

    function completeSimImportRestart() {
        try {
            sessionStorage.setItem(SIM_IMPORT_STATE_KEY, JSON.stringify(state));
            sessionStorage.setItem(SIM_IMPORT_RESTART_FLAG, "1");
        } catch (e) {
            /* private mode */
        }
        location.reload();
    }

    function applyPendingImportRestartState() {
        try {
            if (sessionStorage.getItem(SIM_IMPORT_RESTART_FLAG) !== "1") return;
            var raw = sessionStorage.getItem(SIM_IMPORT_STATE_KEY);
            sessionStorage.removeItem(SIM_IMPORT_RESTART_FLAG);
            sessionStorage.removeItem(SIM_IMPORT_STATE_KEY);
            if (!raw) return;
            var o = JSON.parse(raw);
            if (!o || !o.vessels || !Array.isArray(o.vessels)) return;
            state = Object.assign(state, o);
            pendingImportRestartToast = true;
        } catch (e2) {
            /* ignore */
        }
    }

    function showUacForImportRestart() {
        var bd = document.getElementById("backdropUac");
        if (!bd) {
            completeSimImportRestart();
            return;
        }
        bd.setAttribute("data-uac-context", "import-restart");
        var line1 = document.getElementById("uacHeading");
        var app = bd.querySelector(".uac-app");
        var pub = bd.querySelector(".uac-pub");
        if (line1) line1.textContent = "Do you want to allow this app to make changes to your device?";
        if (app) app.textContent = "eBobServicesController.exe";
        if (pub) pub.textContent = "Verified publisher: BinMaster (simulated)";
        bd.classList.add("show");
        bd.setAttribute("aria-hidden", "false");
    }

    var PRODUCTS = [
        "Portland Cement Type I", "Fly Ash Class C", "Fine Aggregate", "Coarse Aggregate",
        "Flour (wheat)", "Granulated Sugar", "Plastic Pellets", "Wood Pellets",
        "Sodium Hydroxide", "HCl Dilute", "Corn Meal", "Soy Meal",
        "Product A", "Product B", "Product C", "Product D"
    ];
    /** Fill colors — same order as frmVesselSetup BuildColors() */
    var FILL_COLOR_NAMES = [
        "Aqua", "Aquamarine", "Bisque", "Black", "Blanched Almond", "Blue", "Blue Violet", "Brown",
        "Burly Wood", "Cadet Blue", "Chartreuse", "Chocolate", "Coral", "Cornflower Blue", "Crimson", "Cyan",
        "Dark Blue", "Dark Cyan", "Dark Goldenrod", "Dark Gray", "Dark Green", "Dark Khaki", "Dark Magenta",
        "Dark Olive Green", "Dark Orange", "Dark Orchid", "Dark Red", "Dark Salmon", "Dark Sea Green",
        "Dark Slate Blue", "Dark Slate Gray", "Dark Turquoise", "Dark Violet", "Deep Pink", "Deep Sky Blue",
        "Dim Gray", "Dodger Blue", "Firebrick", "Forest Green", "Fuchsia", "Gold", "Goldenrod", "Gray",
        "Green", "Green Yellow", "Hot Pink", "Indian Red", "Indigo", "Khaki", "Lawn Green", "Light Blue",
        "Light Coral", "Light Cyan", "Light Goldenrod Yellow", "Light Green", "Light Pink", "Light Salmon",
        "Light Sea Green", "Light Sky Blue", "Light Slate Gray", "Light Steel Blue", "Lime", "Lime Green",
        "Magenta", "Maroon", "Medium Aquamarine", "Medium Blue", "Medium Orchid", "Medium Purple",
        "Medium Sea Green", "Medium Slate Blue", "Medium Spring Green", "Medium Turquoise", "Medium Violet Red",
        "Midnight Blue", "Misty Rose", "Moccasin", "Navajo White", "Navy", "Olive", "Olive Drab", "Orange",
        "Orange Red", "Orchid", "Pale Goldenrod", "Pale Green", "Pale Turquoise", "Pale Violet Red",
        "Peach Puff", "Peru", "Pink", "Plum", "Powder Blue", "Purple", "Red", "Rosy Brown", "Royal Blue",
        "Saddle Brown", "Salmon", "Sandy Brown", "Sea Green", "Sienna", "Silver", "Sky Blue", "Slate Blue",
        "Slate Gray", "Spring Green", "Steel Blue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet",
        "Wheat", "Yellow", "Yellow Green"
    ];

    var FILL_COLOR_HEX_MAP = {
        Aqua: "#00FFFF", Aquamarine: "#7FFFD4", Bisque: "#FFE4C4", Black: "#000000", "Blanched Almond": "#FFEBCD",
        Blue: "#0000FF", "Blue Violet": "#8A2BE2", Brown: "#A52A2A", "Burly Wood": "#DEB887", "Cadet Blue": "#5F9EA0",
        Chartreuse: "#7FFF00", Chocolate: "#D2691E", Coral: "#FF7F50", "Cornflower Blue": "#6495ED", Crimson: "#DC143C",
        Cyan: "#00FFFF", "Dark Blue": "#00008B", "Dark Cyan": "#008B8B", "Dark Goldenrod": "#B8860B", "Dark Gray": "#A9A9A9",
        "Dark Green": "#006400", "Dark Khaki": "#BDB76B", "Dark Magenta": "#8B008B", "Dark Olive Green": "#556B2F",
        "Dark Orange": "#FF8C00", "Dark Orchid": "#9932CC", "Dark Red": "#8B0000", "Dark Salmon": "#E9967A",
        "Dark Sea Green": "#8FBC8F", "Dark Slate Blue": "#483D8B", "Dark Slate Gray": "#2F4F4F", "Dark Turquoise": "#00CED1",
        "Dark Violet": "#9400D3", "Deep Pink": "#FF1493", "Deep Sky Blue": "#00BFFF", "Dim Gray": "#696969",
        "Dodger Blue": "#1E90FF", Firebrick: "#B22222", "Forest Green": "#228B22", Fuchsia: "#FF00FF", Gold: "#FFD700",
        Goldenrod: "#DAA520", Gray: "#808080", Green: "#008000", "Green Yellow": "#ADFF2F", "Hot Pink": "#FF69B4",
        "Indian Red": "#CD5C5C", Indigo: "#4B0082", Khaki: "#F0E68C", "Lawn Green": "#7CFC00", "Light Blue": "#ADD8E6",
        "Light Coral": "#F08080", "Light Cyan": "#E0FFFF", "Light Goldenrod Yellow": "#FAFAD2", "Light Green": "#90EE90",
        "Light Pink": "#FFB6C1", "Light Salmon": "#FFA07A", "Light Sea Green": "#20B2AA", "Light Sky Blue": "#87CEFA",
        "Light Slate Gray": "#778899", "Light Steel Blue": "#B0C4DE", Lime: "#00FF00", "Lime Green": "#32CD32",
        Magenta: "#FF00FF", Maroon: "#800000", "Medium Aquamarine": "#66CDAA", "Medium Blue": "#0000CD",
        "Medium Orchid": "#BA55D3", "Medium Purple": "#9370DB", "Medium Sea Green": "#3CB371", "Medium Slate Blue": "#7B68EE",
        "Medium Spring Green": "#00FA9A", "Medium Turquoise": "#48D1CC", "Medium Violet Red": "#C71585",
        "Midnight Blue": "#191970", "Misty Rose": "#FFE4E1", Moccasin: "#FFE4B5", "Navajo White": "#FFDEAD", Navy: "#000080",
        Olive: "#808000", "Olive Drab": "#6B8E23", Orange: "#FFA500", "Orange Red": "#FF4500", Orchid: "#DA70D6",
        "Pale Goldenrod": "#EEE8AA", "Pale Green": "#98FB98", "Pale Turquoise": "#AFEEEE", "Pale Violet Red": "#DB7093",
        "Peach Puff": "#FFDAB9", Peru: "#CD853F", Pink: "#FFC0CB", Plum: "#DDA0DD", "Powder Blue": "#B0E0E6", Purple: "#800080",
        Red: "#FF0000", "Rosy Brown": "#BC8F8F", "Royal Blue": "#4169E1", "Saddle Brown": "#8B4513", Salmon: "#FA8072",
        "Sandy Brown": "#F4A460", "Sea Green": "#2E8B57", Sienna: "#A0522D", Silver: "#C0C0C0", "Sky Blue": "#87CEEB",
        "Slate Blue": "#6A5ACD", "Slate Gray": "#708090", "Spring Green": "#00FF7F", "Steel Blue": "#4682B4", Tan: "#D2B48C",
        Teal: "#008080", Thistle: "#D8BFD8", Tomato: "#FF6347", Turquoise: "#40E0D0", Violet: "#EE82EE", Wheat: "#F5DEB3",
        Yellow: "#FFFF00", "Yellow Green": "#9ACD32"
    };

    var VOLUME_DISPLAY_UNITS = [
        "U.S. Gallon (Liquid)", "Cubic Feet", "Cubic Yard", "Imperial Gallon", "Liter", "U.S. Oil Barrel (42 gal)",
        "Cubic Meter", "U.S. Dry Gallon"
    ];

    var WEIGHT_DISPLAY_UNITS = ["Tons", "Pounds", "Kilograms", "Metric Tonnes", "Ounces", "Hundredweight (US)"];

    var DENSITY_UNITS_OPTIONS = [
        "lbs / cubic ft", "lbs / gallon", "kg / cubic meter", "grams / cc", "kg / liter", "tons / cubic yard"
    ];

    /* Combo order matches workstation (VesselTypeID 16 appears before 10–15). */
    var VESSEL_TYPES = [
        { id: 1, name: "Vertical Cylinder" },
        { id: 2, name: "Vertical Cylinder with Cone" },
        { id: 3, name: "Vertical Cylinder with Hemispherical Heads" },
        { id: 4, name: "Vertical Cylinder with Dished Heads" },
        { id: 5, name: "Vertical Cylinder with Ellipsoidal Heads" },
        { id: 6, name: "Horizontal Cylinder" },
        { id: 7, name: "Horizontal Cylinder with Hemispherical Heads" },
        { id: 8, name: "Horizontal Cylinder with Dished Heads" },
        { id: 9, name: "Horizontal Cylinder with Ellipsoidal Heads" },
        { id: 16, name: "Horizontal Cylinder with Conical Heads" },
        { id: 10, name: "Rectangular" },
        { id: 11, name: "Rectangular with Chute" },
        { id: 12, name: "Vertical Oval" },
        { id: 13, name: "Horizontal Oval" },
        { id: 14, name: "Spherical" },
        { id: 15, name: "Custom / Lookup Table" }
    ];

    /* Parameter indices 0–6 = BLL txtParameter1–7 for the selected VesselTypeID. */
    var VESSEL_SHAPE_FIELD_DEFS = {
        1: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Width (W):" }
        ],
        2: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Width (W):" },
            { i: 2, lbl: "Cone Height (CH):" },
            { i: 3, lbl: "Outlet Width (OW):" }
        ],
        3: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Width (W):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        4: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Width (W):" },
            { i: 2, lbl: "Head Height (HH):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        5: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Width (W):" },
            { i: 2, lbl: "Head Height (HH):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        6: [
            { i: 0, lbl: "Length (L):" },
            { i: 1, lbl: "Diameter (D):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        7: [
            { i: 0, lbl: "Length (L):" },
            { i: 1, lbl: "Diameter (D):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        8: [
            { i: 0, lbl: "Length (L):" },
            { i: 1, lbl: "Diameter (D):" },
            { i: 2, lbl: "Head Length (HL):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        9: [
            { i: 0, lbl: "Length (L):" },
            { i: 1, lbl: "Diameter (D):" },
            { i: 2, lbl: "Head Length (HL):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        10: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Length (L):" },
            { i: 2, lbl: "Width (W):" }
        ],
        11: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Length (L):" },
            { i: 2, lbl: "Width (W):" },
            { i: 3, lbl: "Chute Height (CH):" },
            { i: 4, lbl: "Outlet Length (OL):" },
            { i: 5, lbl: "Outlet Width (OW):" }
        ],
        12: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Length (L):" },
            { i: 2, lbl: "Width (W):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        13: [
            { i: 0, lbl: "Height (H):" },
            { i: 1, lbl: "Length (L):" },
            { i: 2, lbl: "Width (W):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        14: [
            { i: 0, lbl: "Diameter (D):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ],
        16: [
            { i: 0, lbl: "Length (L):" },
            { i: 1, lbl: "Diameter (D):" },
            { i: 2, lbl: "Head Length (HL):" },
            { i: 6, lbl: "Full Distance (FD):" }
        ]
    };

    function ensureVesselShapeParams(v) {
        if (!v) return;
        var d = ["25.00", "10.50", "", "", "", "", ""];
        if (!Array.isArray(v.shapeParams) || v.shapeParams.length !== 7) {
            if (v.shapeHeight != null) d[0] = String(v.shapeHeight);
            if (v.shapeWidth != null) d[1] = String(v.shapeWidth);
            v.shapeParams = d;
        } else {
            var j;
            for (j = 0; j < 7; j++) {
                if (v.shapeParams[j] == null) v.shapeParams[j] = "";
            }
        }
        v.shapeHeight = v.shapeParams[0];
        v.shapeWidth = v.shapeParams[1];
    }

    function distanceAbbrevFromHeightLabel(hLabel) {
        var h = hLabel || "";
        return h.indexOf("Meter") >= 0 ? "m" : "ft";
    }

    function volumeAbbrevForDisplay(unitName) {
        var map = {
            "U.S. Gallon (Liquid)": "gal",
            "Cubic Feet": "cu ft",
            "Cubic Yard": "cu yd",
            "Imperial Gallon": "imp gal",
            "Liter": "L",
            "U.S. Oil Barrel (42 gal)": "bbl",
            "Cubic Meter": "m³",
            "U.S. Dry Gallon": "dry gal"
        };
        return map[unitName || ""] || "gal";
    }

    function weightAbbrevForDisplay(unitName) {
        var map = {
            Tons: "tons",
            Pounds: "lbs",
            Kilograms: "kg",
            "Metric Tonnes": "tonne",
            Ounces: "oz",
            "Hundredweight (US)": "cwt"
        };
        return map[unitName || ""] || "tons";
    }

    function customOutputColumnHeader(idx, distAbbr, volAbbr, wtAbbr) {
        if (idx === 0) return "Output";
        if (idx === 1) return "Product Height (" + distAbbr + ")";
        if (idx === 2) return "Product Volume (" + volAbbr + ")";
        if (idx === 3) return "Product Weight (" + wtAbbr + ")";
        return "Output";
    }

    function ensureVesselCustomTable(v) {
        if (!v) return;
        if (!Array.isArray(v.customStrapRows)) {
            v.customStrapRows = [];
        }
        if (v.customOutputTypeIndex == null || v.customOutputTypeIndex === "") {
            v.customOutputTypeIndex = 0;
        }
    }

    function buildVesselCustomTableHtml(v, distAbbr, volAbbr, wtAbbr) {
        ensureVesselCustomTable(v);
        var rows = v.customStrapRows;
        var selIdx = parseInt(v.customOutputTypeIndex, 10);
        if (isNaN(selIdx)) selIdx = 0;
        selIdx = Math.min(3, Math.max(0, selIdx));
        var outHead = customOutputColumnHeader(selIdx, distAbbr, volAbbr, wtAbbr);
        var distHead = "Distance (" + distAbbr + ")";
        var rowParts = rows.map(function (row) {
            var d = row.distance != null ? String(row.distance) : "";
            var o = row.output != null ? String(row.output) : "";
            return (
                '<tr class="vs-custom-data-row">' +
                '<td><input type="text" class="vs-input vs-input-custom vs-custom-d" value="' +
                escapeHtml(d) +
                '"></td>' +
                '<td><input type="text" class="vs-input vs-input-custom vs-custom-o" value="' +
                escapeHtml(o) +
                '"></td>' +
                "</tr>"
            );
        });
        var optLabels = [
            "[SELECT TYPE]",
            "Product Height (" + distAbbr + ")",
            "Product Volume (" + volAbbr + ")",
            "Product Weight (" + wtAbbr + ")"
        ];
        var optHtml = [];
        var oi;
        for (oi = 0; oi < optLabels.length; oi++) {
            optHtml.push(
                '<option value="' +
                oi +
                '"' +
                (oi === selIdx ? " selected" : "") +
                ">" +
                escapeHtml(optLabels[oi]) +
                "</option>"
            );
        }
        return (
            '<div class="vs-custom-panel">' +
            '<div class="vs-custom-toolbar">' +
            '<button type="button" class="vs-custom-btn" id="vs_custom_add">Add Row</button>' +
            '<button type="button" class="vs-custom-btn" id="vs_custom_delete">Delete Row</button>' +
            '<div class="vs-custom-toolbar-spacer" aria-hidden="true"></div>' +
            '<button type="button" class="vs-custom-btn" id="vs_custom_import">Import</button>' +
            '<button type="button" class="vs-custom-btn" id="vs_custom_export">Export</button>' +
            "</div>" +
            '<div class="vs-custom-right-col">' +
            '<div class="vs-custom-grid-outer">' +
            '<table class="vs-custom-grid" id="vs_custom_table">' +
            "<colgroup><col class=\"vs-custom-col-dist\"><col class=\"vs-custom-col-out\"></colgroup>" +
            "<thead><tr>" +
            '<th id="vs_custom_th_dist">' +
            escapeHtml(distHead) +
            "</th>" +
            '<th id="vs_custom_th_out">' +
            escapeHtml(outHead) +
            "</th>" +
            "</tr></thead>" +
            '<tbody id="vs_custom_tbody">' +
            rowParts.join("") +
            "</tbody>" +
            "</table></div>" +
            '<div class="vs-custom-footer">' +
            '<label class="vs-custom-lbl-out" for="vs_custom_output_type">Output Type:</label>' +
            '<select id="vs_custom_output_type" class="vs-select vs-custom-dd" ' +
            'data-dist-abbr="' +
            escapeHtml(distAbbr) +
            '" data-vol-abbr="' +
            escapeHtml(volAbbr) +
            '" data-wt-abbr="' +
            escapeHtml(wtAbbr) +
            '">' +
            optHtml.join("") +
            "</select></div></div></div>"
        );
    }

    function buildVesselShapeDimensionsHtml(v) {
        var tid = parseInt(v.vesselTypeId, 10);
        if (isNaN(tid)) tid = 1;
        if (tid === 15) {
            var hLab = heightUnitsLabel();
            var distAbbr = distanceAbbrevFromHeightLabel(hLab);
            var volAbbr = volumeAbbrevForDisplay(v.volumeDisplayUnits);
            var wtAbbr = weightAbbrevForDisplay(v.weightDisplayUnits);
            return buildVesselCustomTableHtml(v, distAbbr, volAbbr, wtAbbr);
        }
        var shapeParams = v.shapeParams;
        var rows = VESSEL_SHAPE_FIELD_DEFS[tid] || VESSEL_SHAPE_FIELD_DEFS[1];
        var parts = [];
        var r;
        for (r = 0; r < rows.length; r++) {
            var f = rows[r];
            var idx = f.i;
            var raw = shapeParams && shapeParams[idx] != null ? String(shapeParams[idx]) : "";
            parts.push(
                '<div class="vs-ts-dim">' +
                '<label class="vs-ts-dim-lbl" for="vs_sp_' +
                idx +
                '">' +
                escapeHtml(f.lbl) +
                "</label>" +
                '<input type="text" id="vs_sp_' +
                idx +
                '" class="vs-input vs-input-dim vs-shape-param" value="' +
                escapeHtml(raw) +
                '">' +
                "</div>"
            );
        }
        return '<div class="vs-shape-fields">' + parts.join("") + "</div>";
    }

    function readCustomStrapRowsFromDom() {
        var tb = document.getElementById("vs_custom_tbody");
        if (!tb) return null;
        var trs = tb.querySelectorAll("tr.vs-custom-data-row");
        var out = [];
        var i;
        for (i = 0; i < trs.length; i++) {
            var dInp = trs[i].querySelector(".vs-custom-d");
            var oInp = trs[i].querySelector(".vs-custom-o");
            if (dInp && oInp) {
                out.push({ distance: dInp.value, output: oInp.value });
            }
        }
        return out;
    }

    function refreshVesselCustomStrapFromDom(v) {
        if (!v) return;
        var data = readCustomStrapRowsFromDom();
        if (data !== null) {
            v.customStrapRows = data;
        }
        var sel = document.getElementById("vs_custom_output_type");
        if (sel) {
            v.customOutputTypeIndex = parseInt(sel.value, 10);
            if (isNaN(v.customOutputTypeIndex)) v.customOutputTypeIndex = 0;
        }
    }

    function sortCustomStrapTbodyByDistance() {
        var tb = document.getElementById("vs_custom_tbody");
        if (!tb) return;
        var trs = Array.prototype.slice.call(tb.querySelectorAll("tr.vs-custom-data-row"));
        trs.sort(function (a, b) {
            var ad = parseFloat((a.querySelector(".vs-custom-d") || {}).value);
            var bd = parseFloat((b.querySelector(".vs-custom-d") || {}).value);
            if (isNaN(ad)) ad = 0;
            if (isNaN(bd)) bd = 0;
            return ad - bd;
        });
        var i;
        for (i = 0; i < trs.length; i++) {
            tb.appendChild(trs[i]);
        }
    }

    function closeBinventoryMessageBox() {
        var el = document.getElementById("backdropBinventoryMsg");
        if (el) el.classList.remove("show");
    }

    /**
     * WinForms MessageBox-style dialog (frmVesselSetup MsgBox) — stacked above vessel setup modal.
     */
    function showBinventoryMessageBox(opts) {
        opts = opts || {};
        var title = opts.title != null ? opts.title : "Binventory Workstation";
        var message = opts.message || "";
        var icon = opts.icon || "info";
        var buttons = opts.buttons || "ok";
        var el = document.getElementById("backdropBinventoryMsg");
        if (!el) return;
        document.getElementById("binvMsgTitle").textContent = title;
        document.getElementById("binvMsgText").textContent = message;
        var iconEl = document.getElementById("binvMsgIcon");
        iconEl.className = "binv-msg-icon binv-msg-icon-" + icon;
        var footer = document.getElementById("binvMsgFooter");
        if (buttons === "okcancel") {
            footer.innerHTML =
                '<button type="button" class="secondary" id="binvMsgBtnCancel">Cancel</button>' +
                '<button type="button" class="primary" id="binvMsgBtnOk">OK</button>';
        } else {
            footer.innerHTML = '<button type="button" class="primary" id="binvMsgBtnOk">OK</button>';
        }
        function done(result) {
            closeBinventoryMessageBox();
            if (result === "ok" && opts.onOk) opts.onOk();
            if (result === "cancel" && opts.onCancel) opts.onCancel();
        }
        var okBtn = document.getElementById("binvMsgBtnOk");
        var cancelBtn = document.getElementById("binvMsgBtnCancel");
        if (okBtn) okBtn.onclick = function () { done("ok"); };
        if (cancelBtn) cancelBtn.onclick = function () { done("cancel"); };
        var capX = document.getElementById("binvMsgCloseX");
        if (capX) {
            capX.onclick = function () {
                done(buttons === "okcancel" ? "cancel" : "ok");
            };
        }
        el.classList.add("show");
    }

    function parseVsCustomImportText(raw) {
        var lines2 = String(raw || "").split(/\r?\n/);
        var parsed = [];
        var li;
        for (li = 0; li < lines2.length; li++) {
            var line = lines2[li].trim();
            if (!line) continue;
            var parts = line.split(/[\t,]/);
            if (parts.length < 2) continue;
            var a = parts[0].replace(/,/g, "").trim();
            var b = parts[1].replace(/,/g, "").trim();
            if (a === "" || b === "") continue;
            parsed.push({ distance: a, output: b });
        }
        return parsed;
    }

    function applyVsCustomParsedRowsToTable(parsed) {
        var tb3 = document.getElementById("vs_custom_tbody");
        if (!tb3) return;
        tb3.innerHTML = "";
        var pi;
        for (pi = 0; pi < parsed.length; pi++) {
            var tr2 = document.createElement("tr");
            tr2.className = "vs-custom-data-row";
            tr2.innerHTML =
                '<td><input type="text" class="vs-input vs-input-custom vs-custom-d" value="' +
                escapeHtml(parsed[pi].distance) +
                '"></td>' +
                '<td><input type="text" class="vs-input vs-input-custom vs-custom-o" value="' +
                escapeHtml(parsed[pi].output) +
                '"></td>';
            tb3.appendChild(tr2);
        }
        sortCustomStrapTbodyByDistance();
    }

    function refreshVesselShapeParamsFromDom(v, readAsTypeId) {
        if (!v) return;
        var tid =
            readAsTypeId != null
                ? parseInt(readAsTypeId, 10) || 1
                : document.getElementById("vs_vessel_type")
                  ? parseInt(document.getElementById("vs_vessel_type").value, 10) || 1
                  : parseInt(v.vesselTypeId, 10) || 1;
        if (tid === 15) {
            refreshVesselCustomStrapFromDom(v);
            return;
        }
        ensureVesselShapeParams(v);
        var i;
        for (i = 0; i < 7; i++) {
            var el = document.getElementById("vs_sp_" + i);
            if (el) v.shapeParams[i] = el.value;
        }
        v.shapeHeight = v.shapeParams[0];
        v.shapeWidth = v.shapeParams[1];
    }

    /**
     * DeviceTypeID / labels — aligned with frmVesselSetup.vb + BobsBO SensorNetworkModbusRtu.
     * (IDs are not “GWR/Laser/…” by number; workstation loads from DB via GetSupportedSensors.)
     */
    var SENSOR_DEVICE_TYPES = [
        { id: 1, name: "SmartBob-II" },
        { id: 2, name: "SmartBob" },
        { id: 3, name: "SmartBob-II Average" },
        { id: 4, name: "SmartSonic" },
        { id: 5, name: "SmartWave" },
        { id: 6, name: "MNU Ultrasonic" },
        { id: 7, name: "MPX Magnetostrictive" },
        { id: 8, name: "PT-400 Pressure Transmitter" },
        { id: 9, name: "PT-500 Pressure Transmitter" },
        { id: 10, name: "3D Level Scanner" },
        { id: 11, name: "NCR-80 Open Air Radar" },
        { id: 12, name: "GWR-2000 Guided Wave Radar" },
        { id: 13, name: "SmartBob via C-100MB" },
        { id: 14, name: "SPL-100" },
        { id: 15, name: "SPL-200" },
        { id: 16, name: "HART Sensor" }
    ];

    /**
     * At-rest dashboard code 0 — mirrors each sensor class’s GetStatusString(0) in BobsBO.
     */
    function defaultIdleStatusForSensorType(sensorTypeId) {
        var id = parseInt(sensorTypeId, 10);
        if (isNaN(id)) id = 1;
        if (id === 1 || id === 2 || id === 3 || id === 13) return "Retracted";
        if (id === 14) return "Idle";
        if (id === 15) return "Good";
        return "Ready";
    }

    /** Modbus/RTU stack (SensorNetwork RTU) — all use SensorRtuBase.GetStatusString unless overridden (13). */
    function usesRtuBaseStatusString(sensorTypeId) {
        var id = parseInt(sensorTypeId, 10);
        return id >= 4 && id <= 12;
    }

    function isSmartBobC100MB(sensorTypeId) {
        return parseInt(sensorTypeId, 10) === 13;
    }

    /**
     * Mirrors BobsBO SensorBase.GetStatusString (dashboard codes on vessel grid).
     */
    function sensorBaseDashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Retracted";
            case 1:
            case 2:
                return "Descending";
            case 3:
                return "Retracting HT";
            case 4:
            case 13:
                return "Retracting LT";
            case 5:
                return "Manual Retracting HT";
            case 6:
                return "Manual Retracting LT";
            case 7:
            case 8:
                return "Test Cycle Descending";
            case 9:
                return "Test Cycle Retracting HT";
            case 10:
            case 14:
                return "Test Cycle Retracting LT";
            case 11:
                return "Retry Retracting HT";
            case 12:
                return "Retry Retracting LT";
            case 15:
                return "Bob Stuck Top";
            case 16:
                return "Bob Stuck Bottom";
            case 17:
                return "Motor Fault";
            case 18:
                return "Communication Error";
            case 19:
                return "Bob In Override";
            case 20:
                return "Resetting";
            case 55:
                return "Unknown";
            case 56:
                return "Measurement Error";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Mirrors BobsBO SensorRtuBase.GetStatusString (SmartSonic, MNU, MPX, PT-400/500, 3D scanner, NCR-80, GWR-2000, …).
     */
    function sensorRtuDashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Ready";
            case 18:
                return "Communication Error";
            case 55:
                return "Unknown";
            case 57:
                return "Sensor Error";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Mirrors BobsBO Spl100.GetStatusString.
     */
    function spl100DashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Idle";
            case 18:
                return "Communication Error";
            case 52:
                return "Sensor Notification";
            case 53:
                return "Sensor Error";
            case 54:
                return "Parsing Error";
            case 55:
                return "Unknown";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Mirrors BobsBO Spl200.GetStatusString.
     */
    function spl200DashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Good";
            case 18:
                return "Communication Error";
            case 55:
                return "Unknown";
            case 60:
                return "Error";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Mirrors BobsBO HartSensor.GetStatusString.
     */
    function hartSensorDashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Ready";
            case 18:
            case 19:
                return "Communication Error";
            case 52:
                return "Sensor Warning";
            case 53:
                return "Sensor Error";
            case 55:
                return "Unknown";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Mirrors BobsBO SmartBobC100MB.GetStatusString.
     */
    function smartBobC100MBDashboardStatusString(status) {
        switch (status) {
            case 0:
                return "Retracted";
            case 15:
                return "Failed to Drop";
            case 16:
                return "Failed to Retract";
            case 17:
                return "Motor Fault";
            case 18:
                return "Communication Error";
            case 19:
                return "In Override";
            case 20:
                return "Resetting";
            case 55:
                return "Unknown";
            case 56:
                return "Measurement Error";
            case 58:
                return "Measuring";
            case 59:
                return "Retrieving";
            case 90:
                return "Pending";
            default:
                return "Unknown (" + String(status).padStart(2, "0") + ")";
        }
    }

    /**
     * Dispatch by DeviceTypeID — same routing as BobsBO SensorNetwork* + GetSensorStatus().
     */
    function statusStringForDashboard(sensorTypeId, dashboardStatus) {
        var id = parseInt(sensorTypeId, 10);
        if (isNaN(id)) id = 2;
        if (isSmartBobC100MB(id)) return smartBobC100MBDashboardStatusString(dashboardStatus);
        if (id === 14) return spl100DashboardStatusString(dashboardStatus);
        if (id === 15) return spl200DashboardStatusString(dashboardStatus);
        if (id === 16) return hartSensorDashboardStatusString(dashboardStatus);
        if (usesRtuBaseStatusString(id)) return sensorRtuDashboardStatusString(dashboardStatus);
        return sensorBaseDashboardStatusString(dashboardStatus);
    }

    /**
     * Active drop: SmartBob (1–3) “Descending”; C-100MB “Measuring” (58).
     * All other types: Pending → idle (no third label) per RTU/SPL/HART behavior.
     */
    function simulateMiddleMeasurementStatus(sensorTypeId) {
        var id = parseInt(sensorTypeId, 10);
        if (isNaN(id)) id = 2;
        if (usesRtuBaseStatusString(id)) return null;
        if (id === 14 || id === 15 || id === 16) return null;
        if (isSmartBobC100MB(id)) return smartBobC100MBDashboardStatusString(58);
        if (id === 1 || id === 2 || id === 3) return sensorBaseDashboardStatusString(1);
        return null;
    }

    /** SmartBob-II / SmartBob / SmartBob-II Average — BobsBO SmartBobA + SensorBase dashboard codes. */
    function isSmartBobProtocolA(sensorTypeId) {
        var id = parseInt(sensorTypeId, 10);
        return id === 1 || id === 2 || id === 3;
    }

    /**
     * Full Protocol A cycle: Pending → Descending → Retracting HT → Retracting LT → Retracted (BobsBO SensorBase.GetStatusString).
     * Mirrors SmartBobA.SmartBobAState (Drop / RetractHigh / RetractLow) as dashboard codes 1, 3, 4.
     */
    function runSmartBobProtocolAMeasurement(v, opts) {
        opts = opts || {};
        var sid = parseInt(v.sensorTypeId, 10);
        if (isNaN(sid)) sid = 2;
        clearMeasureSimulation(v);
        v._measureTimers = [];
        var timers = v._measureTimers;
        var acc = 0;
        function after(ms, fn) {
            acc += ms;
            timers.push(
                setTimeout(function () {
                    fn();
                }, acc)
            );
        }

        v.status = statusStringForDashboard(sid, 90);
        saveState();
        renderGrid();

        after(280, function () {
            v.status = statusStringForDashboard(sid, 1);
            saveState();
            renderGrid();
        });
        after(2200, function () {
            v.status = statusStringForDashboard(sid, 3);
            saveState();
            renderGrid();
        });
        after(780, function () {
            v.status = statusStringForDashboard(sid, 4);
            saveState();
            renderGrid();
        });
        after(820, function () {
            applySimulatedMeasurementValues(v);
            v.status = defaultIdleStatusForSensorType(sid);
            v.lastMeasurement = new Date().toLocaleString();
            saveState();
            renderGrid();
            v._measureTimers = null;
            if (typeof opts.onDone === "function") opts.onDone();
            if (!opts.silent) toast("Measurement complete — " + v.name);
        });
    }

    /** SmartBob via C-100MB — SmartBobC100MB.vb: Measuring (58) → Retrieving (59) → Retracted (0). */
    function runSmartBobC100MBMeasurement(v, opts) {
        opts = opts || {};
        var sid = 13;
        clearMeasureSimulation(v);
        v._measureTimers = [];
        var timers = v._measureTimers;
        var acc = 0;
        function after(ms, fn) {
            acc += ms;
            timers.push(
                setTimeout(function () {
                    fn();
                }, acc)
            );
        }

        v.status = statusStringForDashboard(sid, 90);
        saveState();
        renderGrid();

        after(260, function () {
            v.status = statusStringForDashboard(sid, 58);
            saveState();
            renderGrid();
        });
        after(2000, function () {
            v.status = statusStringForDashboard(sid, 59);
            saveState();
            renderGrid();
        });
        after(720, function () {
            applySimulatedMeasurementValues(v);
            v.status = defaultIdleStatusForSensorType(sid);
            v.lastMeasurement = new Date().toLocaleString();
            saveState();
            renderGrid();
            v._measureTimers = null;
            if (typeof opts.onDone === "function") opts.onDone();
            if (!opts.silent) toast("Measurement complete — " + v.name);
        });
    }

    function runGenericMeasurement(v, opts) {
        opts = opts || {};
        clearMeasureSimulation(v);
        var sid = parseInt(v.sensorTypeId, 10);
        if (isNaN(sid)) sid = 2;

        v.status = statusStringForDashboard(sid, 90);
        saveState();
        renderGrid();

        var mid = simulateMiddleMeasurementStatus(sid);
        var finish = function () {
            applySimulatedMeasurementValues(v);
            v.status = defaultIdleStatusForSensorType(sid);
            v.lastMeasurement = new Date().toLocaleString();
            saveState();
            renderGrid();
            v._measureTimers = null;
            if (typeof opts.onDone === "function") opts.onDone();
            if (!opts.silent) toast("Measurement complete — " + v.name);
        };

        if (mid === null) {
            v._measureTimers = [setTimeout(finish, 450)];
            return;
        }

        v._measureTimers.push(
            setTimeout(function () {
                v.status = mid;
                saveState();
                renderGrid();
            }, 280)
        );
        v._measureTimers.push(setTimeout(finish, 750));
    }

    function clearMeasureSimulation(v) {
        if (v._measureTimers && v._measureTimers.length) {
            v._measureTimers.forEach(function (tid) {
                clearTimeout(tid);
            });
        }
        v._measureTimers = [];
    }

    var VESSEL_NAMES = [
        "Cement Silo 1", "Cement Silo 2", "Fly Ash East", "Fly Ash West",
        "Sand Bin A", "Sand Bin B", "Aggregate #3", "Aggregate #4",
        "Chemical Tank 1", "Chemical Tank 2", "Flour Silo", "Sugar Silo",
        "Pellet Day", "Pellet Night", "Reserve 1", "Reserve 2",
        "Overflow A", "Overflow B", "Staging N", "Staging S",
        "Mill Feed", "Coarse Bin", "Fine Bin", "Dust Collector",
        "Silo 25", "Silo 26", "Silo 27", "Silo 28",
        "Tank Cold", "Tank Hot", "Mixer Upper", "Mixer Lower"
    ];

    var state = {
        currentUser: "Admin",
        currentPage: 0,
        vessels: [],
        groups: [],
        schedules: [],
        /** Ordered vessel ids for dashboard when temp group is applied (mirrors gTempGroupVessels + ApplyTempGroup). null = full site list. */
        tempGroupVesselIds: null,
        contacts: [],
        sites: [],
        sensorNetworks: [],
        /** frmEmailSetup — tutorial defaults match typical eBob email setup (Enable SMTP on; port 25; sample addresses/subjects). */
        emailSettings: {
            enableSmtpClient: true,
            smtp: "",
            port: "25",
            useSecureConnection: false,
            useSmtpAuth: false,
            smtpUserId: "",
            smtpPassword: "",
            smtpVerifyPassword: "",
            adminEmail: "admin@ebob.com",
            defaultFrom: "messenger@ebob.com",
            fromAddr: "messenger@ebob.com",
            adminSubject: "Administrative Alert",
            measureSubject: "Vessel Status",
            alarmSubject: "Vessel Alarm",
            emailAppend: ""
        },
        systemSettings: {
            units: "Imperial",
            timezone: "US/Central",
            registeredUser: "",
            companyName: "",
            streetAddress: "",
            streetAddress2: "",
            city: "",
            state: "",
            zipCode: "",
            country: "",
            ldapAddress: "",
            measurementRetentionDays: "60",
            autoLogin: true
        },
        users: [
            {
                id: "u1",
                userId: "Admin",
                name: "Admin",
                role: "Administrator",
                userType: "Administrator",
                firstName: "Admin",
                middleInit: "",
                lastName: "",
                lastLogon: ""
            }
        ],
        emailReportPrefs: { to: "operator@example.com", frequency: "Daily" },
        currentWorkstationSiteId: "st1",
        /** eBob Engine / Scheduler up — BobMsgQue reachable (Vessel.vb GetSensorStatus). */
        ebobServicesRunning: true,
        /**
         * Mirrors AppGlobals.vb gbVesselsReadOnly — set True when engine connection is lost (VesselUtility.vb
         * after GetSensorStatus fails); not cleared when services start again until Binventory restarts (full reload / factory reset).
         */
        vesselsReadOnly: false
    };

    function escapeHtml(s) {
        if (s == null) return "";
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function uid(prefix) {
        return prefix + "-" + Math.random().toString(36).slice(2, 9);
    }

    function seedVessel(i) {
        var pct = 40 + (i * 7) % 55;
        var vol = Math.round(1000 + i * 33);
        var prod = PRODUCTS[i % PRODUCTS.length];
        return {
            id: "v" + (i + 1),
            vesselNumericId: i + 1,
            sortOrder: i + 1,
            sensorNetworkId: "n1",
            name: VESSEL_NAMES[i % VESSEL_NAMES.length],
            product: prod,
            contents: prod,
            pctFull: pct,
            heightFt: 4 + (pct / 100) * 8,
            volumeCuFt: vol,
            weightLb: Math.round(vol * 40),
            status: "Unknown",
            lastMeasurement: new Date().toLocaleString(),
            headroom: false,
            capacityHeightFt: 14,
            siteId: null,
            fillColor: "Light Blue",
            volumeDisplayUnits: "U.S. Gallon (Liquid)",
            weightDisplayUnits: "Tons",
            defaultHeadroom: false,
            vesselTypeId: 1,
            sensorTypeId: 1,
            densityMode: "density",
            productDensity: "40",
            densityUnits: "lbs / cubic ft",
            specificGravity: "1.00",
            alarmPreLowEnabled: false,
            alarmPreLowPct: "",
            alarmLowEnabled: false,
            alarmLowPct: "",
            alarmHighEnabled: false,
            alarmHighPct: "",
            alarmPreHighEnabled: false,
            alarmPreHighPct: "",
            emailNotificationsEnabled: false,
            emailFlags: { high: true, preHigh: true, preLow: true, low: true, vesselStatus: false, error: false },
            vesselContactIds: [],
            shapeHeight: (25 + (i % 7) * 0.25).toFixed(2),
            shapeWidth: (10.5 + (i % 5) * 0.15).toFixed(2),
            shapeParams: [
                (25 + (i % 7) * 0.25).toFixed(2),
                (10.5 + (i % 5) * 0.15).toFixed(2),
                "",
                "",
                "",
                "",
                ""
            ],
            customStrapRows: [],
            customOutputTypeIndex: 0,
            verticalSplitPartitioned: false,
            partitionScale: ""
        };
    }

    /** Sites + contacts only — vessel/network data is loaded per workstation site (reload). */
    function seedSitesAndContacts() {
        state.contacts = [
            { id: "c1", name: "Jane Operator", email: "jane@example.com", phone: "555-0100" },
            { id: "c2", name: "John Smith", email: "john@example.com", phone: "555-0102" }
        ];
        state.sites = [
            {
                id: "st1",
                name: "Current Location",
                workstationSiteId: 1,
                serviceHostIp: "127.0.0.1",
                serviceHostPort: "8093",
                companyName: "Acme Manufacturing",
                streetAddress: "100 Industrial Pkwy",
                streetAddress2: "",
                city: "Minneapolis",
                state: "MN",
                zip: "55401",
                country: "USA",
                distanceUnitsId: "1"
            },
            {
                id: "st2",
                name: "Warehouse",
                workstationSiteId: 2,
                serviceHostIp: "10.0.0.12",
                serviceHostPort: "8093",
                companyName: "Acme Logistics",
                streetAddress: "2200 Warehouse Rd",
                streetAddress2: "Dock B",
                city: "St. Paul",
                state: "MN",
                zip: "55101",
                country: "USA",
                distanceUnitsId: "2"
            }
        ];
    }

    /** Full factory defaults — same as first visit (nothing from localStorage). */
    function resetStateToFactoryDefaults() {
        state.currentUser = "Admin";
        state.currentPage = 0;
        state.emailSettings = {
            enableSmtpClient: true,
            smtp: "",
            port: "25",
            useSecureConnection: false,
            useSmtpAuth: false,
            smtpUserId: "",
            smtpPassword: "",
            smtpVerifyPassword: "",
            adminEmail: "admin@ebob.com",
            defaultFrom: "messenger@ebob.com",
            fromAddr: "messenger@ebob.com",
            adminSubject: "Administrative Alert",
            measureSubject: "Vessel Status",
            alarmSubject: "Vessel Alarm",
            emailAppend: ""
        };
        state.systemSettings = {
            units: "Imperial",
            timezone: "US/Central",
            registeredUser: "",
            companyName: "",
            streetAddress: "",
            streetAddress2: "",
            city: "",
            state: "",
            zipCode: "",
            country: "",
            ldapAddress: "",
            measurementRetentionDays: "60",
            autoLogin: true
        };
        state.users = [
            {
                id: "u1",
                userId: "Admin",
                name: "Admin",
                role: "Administrator",
                userType: "Administrator",
                firstName: "Admin",
                middleInit: "",
                lastName: "",
                lastLogon: ""
            }
        ];
        state.emailReportPrefs = { to: "operator@example.com", frequency: "Daily" };
        state.currentWorkstationSiteId = "st1";
        state.ebobServicesRunning = true;
        state.vesselsReadOnly = false;
        seedSitesAndContacts();
        state.sites.forEach(ensureSiteFields);
        reloadWorkstationForSite(state.currentWorkstationSiteId, { skipRefreshUI: true });
        simSysExportDatMeta = null;
        simSysExportSnapshotJson = null;
    }

    function ensureSiteFields(s) {
        if (!s) return;
        if (s.workstationSiteId == null) {
            var m = /^st(\d+)$/.exec(s.id || "");
            s.workstationSiteId = m ? parseInt(m[1], 10) : 1;
        }
        if (s.serviceHostIp == null) s.serviceHostIp = "";
        if (s.serviceHostPort == null) s.serviceHostPort = "";
        if (s.companyName == null) s.companyName = "";
        if (s.streetAddress == null) s.streetAddress = "";
        if (s.streetAddress2 == null) s.streetAddress2 = "";
        if (s.city == null) s.city = "";
        if (s.state == null) s.state = "";
        if (s.zip == null) s.zip = "";
        if (s.country == null) s.country = "";
        if (s.distanceUnitsId == null) s.distanceUnitsId = "1";
    }

    function ensureVesselFields(v, index) {
        var i = index != null ? index : 0;
        if (v.sortOrder == null) v.sortOrder = i + 1;
        if (v.sensorNetworkId == null && state.sensorNetworks && state.sensorNetworks[0]) {
            v.sensorNetworkId = state.sensorNetworks[0].id;
        }
        if (v.vesselNumericId == null) {
            var m = /^v(\d+)$/.exec(v.id || "");
            v.vesselNumericId = m ? parseInt(m[1], 10) : i + 1;
        }
        if (v.contents == null && v.product) v.contents = v.product;
        ensureVesselSetupDefaults(v);
        if (v.status === "Idle" || v.status === "OK") {
            v.status = defaultIdleStatusForSensorType(v.sensorTypeId);
        }
        if (!vesselHasValidSavedSensorAddress(v)) {
            assignUniqueSensorAddressForVessel(v);
        }
    }

    function ensureVesselSetupDefaults(v) {
        if (v.fillColor == null) v.fillColor = "Light Blue";
        if (v.volumeDisplayUnits == null) v.volumeDisplayUnits = "U.S. Gallon (Liquid)";
        if (v.weightDisplayUnits == null) v.weightDisplayUnits = "Tons";
        if (v.defaultHeadroom == null) v.defaultHeadroom = !!v.headroom;
        if (v.headroom == null) v.headroom = !!v.defaultHeadroom;
        if (v.capacityHeightFt == null) v.capacityHeightFt = 14;
        if (v.vesselTypeId == null) v.vesselTypeId = 1;
        if (v.sensorTypeId == null) v.sensorTypeId = 2;
        if (v.densityMode == null) v.densityMode = "density";
        if (v.productDensity == null) v.productDensity = "40";
        if (v.densityUnits == null) v.densityUnits = "lbs / cubic ft";
        if (v.specificGravity == null) v.specificGravity = "1.00";
        if (v.alarmPreLowEnabled == null) v.alarmPreLowEnabled = false;
        if (v.alarmPreLowPct == null) v.alarmPreLowPct = "";
        if (v.alarmLowEnabled == null) v.alarmLowEnabled = false;
        if (v.alarmLowPct == null) v.alarmLowPct = "";
        if (v.alarmHighEnabled == null) v.alarmHighEnabled = false;
        if (v.alarmHighPct == null) v.alarmHighPct = "";
        if (v.alarmPreHighEnabled == null) v.alarmPreHighEnabled = false;
        if (v.alarmPreHighPct == null) v.alarmPreHighPct = "";
        if (v.emailNotificationsEnabled == null) v.emailNotificationsEnabled = false;
        if (!v.emailFlags) {
            v.emailFlags = { high: true, preHigh: true, preLow: true, low: true, vesselStatus: false, error: false };
        }
        if (v.shapeHeight == null) v.shapeHeight = "25.00";
        if (v.shapeWidth == null) v.shapeWidth = "10.50";
        ensureVesselShapeParams(v);
        ensureVesselCustomTable(v);
        if (v.verticalSplitPartitioned == null) v.verticalSplitPartitioned = false;
        if (v.partitionScale == null) v.partitionScale = "";
        if (v.sensorEnabled == null) v.sensorEnabled = true;
        if (v.sensorAddress == null) v.sensorAddress = "";
        if (v.sensorOffset == null) v.sensorOffset = "";
        if (v.sensorDecimalPlaces == null) v.sensorDecimalPlaces = "2";
        if (v.sensorDistanceVariable == null) v.sensorDistanceVariable = "SV";
        var dpNorm = String(v.sensorDecimalPlaces);
        if (dpNorm !== "1" && dpNorm !== "2") v.sensorDecimalPlaces = "2";
        var dvNorm = String(v.sensorDistanceVariable || "").toUpperCase();
        if (["PV", "SV", "TV", "QV"].indexOf(dvNorm) < 0) v.sensorDistanceVariable = "SV";
        if (v.vegaSensorCount == null) v.vegaSensorCount = "1";
        if (v.sbEnableMaxDrop == null) v.sbEnableMaxDrop = false;
        if (v.sbMaxDrop == null) v.sbMaxDrop = "";
        var net = getSensorNetworkById(v.sensorNetworkId);
        if (net) {
            var allowedSt = sensorTypeIdsForNetwork(net);
            var stSeen = {};
            for (var si = 0; si < allowedSt.length; si++) {
                stSeen[allowedSt[si]] = true;
            }
            var curSt = parseInt(v.sensorTypeId, 10);
            if (isNaN(curSt) || !stSeen[curSt]) {
                v.sensorTypeId = allowedSt[0] != null ? allowedSt[0] : 1;
            }
        }
        ensureVegaSensorsArray(v);
        ensureVesselContactIds(v);
    }

    /** frmVesselSetup: pnlVegaSensor — per-row Enable, Address, Offset, Distance Variable (PV/SV/TV/QV). */
    var VEGA_DISTANCE_VARIABLES = ["PV", "SV", "TV", "QV"];

    function ensureVegaSensorsArray(v) {
        if (!v) return;
        if (!Array.isArray(v.vegaSensors)) v.vegaSensors = [];
        var i;
        for (i = 0; i < 32; i++) {
            if (!v.vegaSensors[i] || typeof v.vegaSensors[i] !== "object") {
                v.vegaSensors[i] = {
                    enabled: i === 0,
                    address: String(i + 1),
                    offset: i === 0 ? "0.00" : "",
                    dv: "SV"
                };
            } else {
                var row = v.vegaSensors[i];
                if (row.enabled == null) row.enabled = i === 0;
                if (row.address == null) row.address = String(i + 1);
                if (row.offset == null) row.offset = i === 0 ? "0.00" : "";
                var dv = String(row.dv || row.distanceVariable || "SV").toUpperCase();
                if (VEGA_DISTANCE_VARIABLES.indexOf(dv) < 0) dv = "SV";
                row.dv = dv;
            }
        }
    }

    /** eBob BLL: sensor address min/max by tblSensorType (Case Else + Vega). */
    function getSensorAddressBounds(sensorTypeId) {
        var st = parseInt(sensorTypeId, 10) || 1;
        switch (st) {
            case 8:
            case 9:
                return { min: 1, max: 247 };
            case 10:
                return { min: 1, max: 64 };
            case 13:
                return { min: 1, max: 120 };
            case 16:
                return { min: 0, max: 63 };
            default:
                return { min: 1, max: 255 };
        }
    }

    function validateIntegerSensorAddress(raw, min, max) {
        var t = String(raw == null ? "" : raw).trim();
        if (t === "") {
            return {
                ok: false,
                msg:
                    "Sensor Address is required and must be a whole number ranging from " +
                    min +
                    " to " +
                    max +
                    "."
            };
        }
        if (!/^\d+$/.test(t)) {
            return {
                ok: false,
                msg:
                    "Sensor Address is required and must be a whole number ranging from " +
                    min +
                    " to " +
                    max +
                    "."
            };
        }
        var n = parseInt(t, 10);
        if (n < min || n > max) {
            return {
                ok: false,
                msg:
                    "Sensor Address is required and must be a whole number ranging from " +
                    min +
                    " to " +
                    max +
                    "."
            };
        }
        return { ok: true, normalized: String(n) };
    }

    function vesselDeviceAddressStrings(v) {
        if (!v) return [];
        ensureVesselSetupDefaults(v);
        var st = parseInt(v.sensorTypeId, 10) || 1;
        var out = [];
        if (st === 11 || st === 12) {
            ensureVegaSensorsArray(v);
            var n = Math.min(32, Math.max(1, parseInt(v.vegaSensorCount || "1", 10) || 1));
            var i;
            var b = getSensorAddressBounds(st);
            for (i = 0; i < n; i++) {
                var row = v.vegaSensors[i];
                var raw = row && row.address != null ? row.address : "";
                var vr = validateIntegerSensorAddress(raw, b.min, b.max);
                if (vr.ok) out.push(vr.normalized);
            }
        } else {
            var b2 = getSensorAddressBounds(st);
            var vr2 = validateIntegerSensorAddress(v.sensorAddress, b2.min, b2.max);
            if (vr2.ok) out.push(vr2.normalized);
        }
        return out;
    }

    function collectUsedSensorAddressesOnNetwork(networkId, excludeVesselId) {
        var used = {};
        state.vessels.forEach(function (ov) {
            if (ov.id === excludeVesselId) return;
            if (ov.sensorNetworkId !== networkId) return;
            vesselDeviceAddressStrings(ov).forEach(function (a) {
                used[a] = true;
            });
        });
        return used;
    }

    function firstFreeSensorAddress(usedMap) {
        var p;
        for (p = 1; p <= 255; p++) {
            var s = String(p);
            if (!usedMap[s]) return s;
        }
        return "1";
    }

    /** Every vessel must have valid device address(es); mirrors eBob tblDevice insert + BLL validation. */
    function assignUniqueSensorAddressForVessel(nv) {
        if (!nv) return;
        ensureVesselSetupDefaults(nv);
        var netId = nv.sensorNetworkId;
        if (!netId) return;
        var used = collectUsedSensorAddressesOnNetwork(netId, nv.id);
        var st = parseInt(nv.sensorTypeId, 10) || 1;
        if (st === 11 || st === 12) {
            ensureVegaSensorsArray(nv);
            var n = Math.min(32, Math.max(1, parseInt(nv.vegaSensorCount || "1", 10) || 1));
            var i;
            for (i = 0; i < n; i++) {
                var addr = firstFreeSensorAddress(used);
                nv.vegaSensors[i].address = addr;
                used[addr] = true;
            }
            nv.sensorAddress = nv.vegaSensors[0] ? nv.vegaSensors[0].address : "1";
        } else {
            nv.sensorAddress = firstFreeSensorAddress(used);
        }
    }

    function vesselHasValidSavedSensorAddress(v) {
        if (!v) return false;
        ensureVesselSetupDefaults(v);
        var st = parseInt(v.sensorTypeId, 10) || 1;
        var b = getSensorAddressBounds(st);
        if (st === 11 || st === 12) {
            ensureVegaSensorsArray(v);
            var n = Math.min(32, Math.max(1, parseInt(v.vegaSensorCount || "1", 10) || 1));
            var i;
            for (i = 0; i < n; i++) {
                var ra = v.vegaSensors[i] && v.vegaSensors[i].address;
                if (!validateIntegerSensorAddress(ra, b.min, b.max).ok) return false;
            }
            return true;
        }
        return validateIntegerSensorAddress(v.sensorAddress, b.min, b.max).ok;
    }

    /** Read frmVesselSetup sensor tab; return eBob-style errors before mutating vessel state. */
    function validateVesselSetupSensorBlockFromDom() {
        var vid = vmSetupEditingId;
        var netEl = document.getElementById("vs_network");
        var stEl = document.getElementById("vs_sensor_type");
        var networkId = netEl ? netEl.value : "";
        if (!networkId || String(networkId).trim() === "") {
            return { ok: false, message: "A valid sensor network must be selected." };
        }
        var st = parseInt(stEl && stEl.value, 10) || 0;
        if (st < 1) {
            return { ok: false, message: "A sensor type must be selected." };
        }
        var b = getSensorAddressBounds(st);
        var pending = [];
        var dupMsg = "There are duplicate sensor addresses. Sensor addresses must be unique.";
        if (st === 11 || st === 12) {
            var vcEl = document.getElementById("vs_vega_count");
            var n = Math.min(32, Math.max(1, parseInt(vcEl && vcEl.value, 10) || 1));
            var seen = {};
            var i;
            for (i = 1; i <= n; i++) {
                var aEl = document.getElementById("vs_vg_a" + i);
                var raw = aEl ? aEl.value : "";
                var vr = validateIntegerSensorAddress(raw, b.min, b.max);
                if (!vr.ok) {
                    return {
                        ok: false,
                        message:
                            "Sensor addresses must be a whole number ranging from " + b.min + " to " + b.max + "."
                    };
                }
                if (seen[vr.normalized]) return { ok: false, message: dupMsg };
                seen[vr.normalized] = true;
                pending.push(vr.normalized);
            }
        } else if (st === 1 || st === 3) {
            var sbAddr = document.getElementById("vs_sb_addr1");
            var raw = sbAddr ? sbAddr.value : "";
            var vr = validateIntegerSensorAddress(raw, b.min, b.max);
            if (!vr.ok) return { ok: false, message: vr.msg };
            pending.push(vr.normalized);
        } else {
            var genAddr = document.getElementById("vs_sensor_address");
            var rawG = genAddr ? genAddr.value : "";
            var vrG = validateIntegerSensorAddress(rawG, b.min, b.max);
            if (!vrG.ok) return { ok: false, message: vrG.msg };
            pending.push(vrG.normalized);
        }
        var pi;
        for (pi = 0; pi < pending.length; pi++) {
            var other = findVesselWithSensorAddressOnNetwork(networkId, vid, pending[pi]);
            if (other) {
                return {
                    ok: false,
                    message:
                        "A sensor with this address is already assigned to '" +
                        (other.name || "Unknown") +
                        "' on the same sensor network."
                };
            }
        }
        return { ok: true };
    }

    function findVesselWithSensorAddressOnNetwork(networkId, excludeVesselId, normalizedAddr) {
        var vi;
        for (vi = 0; vi < state.vessels.length; vi++) {
            var ov = state.vessels[vi];
            if (ov.id === excludeVesselId) continue;
            if (ov.sensorNetworkId !== networkId) continue;
            var addrs = vesselDeviceAddressStrings(ov);
            var ai;
            for (ai = 0; ai < addrs.length; ai++) {
                if (addrs[ai] === normalizedAddr) return ov;
            }
        }
        return null;
    }

    function vegaOffsetHeaderText() {
        return state.systemSettings && state.systemSettings.units === "Metric" ? "Offset (m)" : "Offset (ft)";
    }

    function buildVegaDistanceVariableOptionsHtml(selectedDv) {
        var dvSel = String(selectedDv || "SV").toUpperCase();
        if (VEGA_DISTANCE_VARIABLES.indexOf(dvSel) < 0) dvSel = "SV";
        return VEGA_DISTANCE_VARIABLES.map(function (L) {
            return (
                '<option value="' +
                L +
                '"' +
                (dvSel === L ? " selected" : "") +
                ">" +
                L +
                "</option>"
            );
        }).join("");
    }

    function buildVegaSensorTbodyRowsHtml(v, count) {
        ensureVegaSensorsArray(v);
        var n = Math.min(32, Math.max(1, count || 1));
        var parts = [];
        var i;
        for (i = 0; i < n; i++) {
            var row = v.vegaSensors[i];
            var idx1 = i + 1;
            var en = row.enabled === true;
            var addr = escapeHtml(String(row.address != null ? row.address : ""));
            var off = escapeHtml(String(row.offset != null ? row.offset : ""));
            var dopts = buildVegaDistanceVariableOptionsHtml(row.dv);
            parts.push(
                '<tr>' +
                    '<td class="vs-vega-td-chk"><input type="checkbox" id="vs_vg_en' +
                    idx1 +
                    '"' +
                    (en ? " checked" : "") +
                    "></td>" +
                    '<td><input type="text" class="vs-input vs-input-sm" id="vs_vg_a' +
                    idx1 +
                    '" value="' +
                    addr +
                    '"' +
                    (en ? "" : " disabled") +
                    "></td>" +
                    '<td><input type="text" class="vs-input vs-input-sm" id="vs_vg_o' +
                    idx1 +
                    '" value="' +
                    off +
                    '"' +
                    (en ? "" : " disabled") +
                    "></td>" +
                    '<td><select id="vs_vg_dv' +
                    idx1 +
                    '" class="vs-select vs-select-vega-dv"' +
                    (en ? "" : " disabled") +
                    ">" +
                    dopts +
                    "</select></td>" +
                    "</tr>"
            );
        }
        return parts.join("");
    }

    /** Product vs headroom height/volume/weight — matches Vessel.vb RefreshControl chkHeadroomDisplay. */
    function vesselDisplayMetrics(v) {
        var pct = Math.max(0, Math.min(100, Number(v.pctFull) || 0));
        var capH = v.capacityHeightFt != null ? Number(v.capacityHeightFt) : 14;
        if (isNaN(capH) || capH <= 0) capH = 14;
        var prodH = (pct / 100) * capH;
        var headH = Math.max(0, capH - prodH);
        var prodVol = Number(v.volumeCuFt) || 0;
        var capVol = pct > 0.05 ? prodVol / (pct / 100) : prodVol;
        var headVol = Math.max(0, capVol - prodVol);
        var prodWt = Number(v.weightLb) || 0;
        var capWt = pct > 0.05 ? prodWt / (pct / 100) : prodWt;
        var headWt = Math.max(0, capWt - prodWt);
        return {
            product: {
                heightFt: prodH,
                volumeCuFt: Math.round(prodVol),
                weightLb: Math.round(prodWt)
            },
            headroom: {
                heightFt: headH,
                volumeCuFt: Math.round(headVol),
                weightLb: Math.round(headWt)
            }
        };
    }

    function heightUnitsLabel() {
        return state.systemSettings && state.systemSettings.units === "Metric" ? "Meters (m)" : "Feet (ft)";
    }

    function fillColorPreviewCss(name) {
        var hx = FILL_COLOR_HEX_MAP[name] || "#cccccc";
        return "linear-gradient(180deg, " + hx + " 0%, #ffffff 100%)";
    }

    var VS_ASSIST_DEFAULT = "<<< Place the mouse cursor over a control to get assistance. >>>";

    function assistanceTextForVesselType(id) {
        var m = {
            1: "This is an upright cylinder with a flat bottom. This may be used for a grain bin or a liquid tank.",
            2: "This is an upright cylinder with a hopper or cone bottom. This may be used for a hopper bin.",
            3: "This is a vertical capsule or upright cylinder with hemispherical heads.",
            4: "This is an upright cylinder with dished, bumped or spherical heads.",
            5: "This is an upright cylinder with ellipsoidal heads. It may be used on tanks with ASME or DIN standard heads, see the operators manual for details.",
            6: "This is a horizontal cylinder with flat heads.",
            7: "This is a horizontal capsule or cylinder with hemispherical heads.",
            8: "This is a horizontal cylinder with dished, bumped or spherical heads.",
            9: "This is a horizontal cylinder with ellipsoidal heads. It may be used on tanks with ASME or DIN standard heads, see the operators manual for details.",
            10: "This is a cube or rectangular vessel.",
            11: "This is a rectangular vessel with a hopper or chute on the bottom.",
            12: "This is an upright oval vessel.",
            13: "This is an oval vessel laying flat.",
            14: "This is a spherical vessel.",
            15: "Custom / Lookup Table — define the vessel profile using the custom distance/volume table in eBob.",
            16: "This is a horizontal cylinder with conical heads."
        };
        return m[id] || "";
    }

    /**
     * Vessel Type preview — same bitmaps as frmVesselSetup picVesselType (My.Resources.*).
     * PNGs copied from eBobWorkstation/Resources into assets/vessel-types/.
     */
    var VESSEL_TYPE_IMAGE_BASE = "assets/vessel-types/";

    var VESSEL_TYPE_IMAGE_FILE = {
        1: "01_VerticalCylinder_.PNG",
        2: "02_VerticalCylinderWithCone_.PNG",
        3: "03_VerticalCylinderWithHemisphericalHeads_.PNG",
        4: "04_VerticalCylinderWithDishedHeads_.PNG",
        5: "05_VerticalCylinderWithEllipsoidalHeads_.PNG",
        6: "06_HorizontalCylinder_.PNG",
        7: "07_HorizontalCylinderWithHemisphericalHeads_.PNG",
        8: "08_HorizontalCylinderWithDishedHeads_.PNG",
        9: "09_HorizontalCylinderWithEllipsoidalHeads_.PNG",
        10: "10_Rectangular_.PNG",
        11: "11_RectangularWithChute_.PNG",
        12: "12_VerticalOval_.PNG",
        13: "13_HorizontalOval_.PNG",
        14: "14_Spherical_.PNG",
        16: "16_HorizontalCylinderWithConicalHeads_.PNG"
    };

    /** Type 15 Custom — workstation sets picVesselType.Image = Nothing; placeholder SVG. */
    function buildVesselShapeSvgCustomOnly() {
        var stroke = "#000";
        var sw = 1.5;
        return (
            '<svg class="vs-shape-svg" viewBox="0 0 278 278" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' +
            '<rect width="278" height="278" fill="#fff"/>' +
            '<g stroke="' +
            stroke +
            '" stroke-width="' +
            sw +
            '" fill="none">' +
            '<rect x="40" y="60" width="198" height="140"/>' +
            '<line x1="40" y1="90" x2="238" y2="90"/>' +
            '<line x1="40" y1="120" x2="238" y2="120"/>' +
            '<line x1="40" y1="150" x2="238" y2="150"/>' +
            '<line x1="90" y1="60" x2="90" y2="200"/>' +
            '<line x1="140" y1="60" x2="140" y2="200"/>' +
            '<line x1="190" y1="60" x2="190" y2="200"/>' +
            "</g>" +
            '<text x="70" y="240" font-family="Tahoma,Arial,sans-serif" font-size="11" fill="#555">Custom profile table</text>' +
            "</svg>"
        );
    }

    /** frmVesselSetup cboVesselType_SelectedIndexChanged: pnlSplitVessel only for types 1, 2, 10, 11 */
    function vesselTypeShowsVerticalSplitPartition(typeId) {
        var tid = parseInt(typeId, 10);
        if (isNaN(tid)) tid = 1;
        return tid === 1 || tid === 2 || tid === 10 || tid === 11;
    }

    function buildVesselShapePreviewMarkup(typeId, dimH, dimW) {
        var tid = parseInt(typeId, 10);
        if (isNaN(tid)) tid = 1;
        if (tid === 15) {
            return buildVesselShapeSvgCustomOnly();
        }
        var fn = VESSEL_TYPE_IMAGE_FILE[tid] || VESSEL_TYPE_IMAGE_FILE[1];
        var src = VESSEL_TYPE_IMAGE_BASE + fn;
        return (
            '<img class="vs-vessel-type-img" src="' +
            escapeHtml(src) +
            '" width="278" height="278" alt="" draggable="false">'
        );
    }

    function loadState() {
        try {
            localStorage.removeItem(STORAGE_KEY);
        } catch (e) {
            /* private mode / blocked storage */
        }
        resetStateToFactoryDefaults();
        applyPendingImportRestartState();
    }

    function saveState() {
        /* Session-only: state is not persisted. Each full page load starts from factory defaults. */
    }

    function toast(msg) {
        var el = document.getElementById("toast");
        el.textContent = msg;
        el.classList.add("show");
        clearTimeout(toast._t);
        toast._t = setTimeout(function () { el.classList.remove("show"); }, 2400);
    }

    function updateTitleBar() {
        var t = document.querySelector(".title-bar .title");
        var siteId = state.currentWorkstationSiteId;
        var site =
            state.sites && siteId
                ? state.sites.filter(function (s) {
                      return s.id === siteId;
                  })[0]
                : null;
        var siteDesc = site && site.name ? site.name : "Current Location";
        t.textContent = "Binventory Workstation - Site: " + siteDesc;
    }

    /** True when gsLoggedInUserID equivalent is set (frmMain after login). */
    function isSessionLoggedIn() {
        return state.currentUser != null && String(state.currentUser).trim() !== "";
    }

    /**
     * frmMain.LoggedInMenuDefaults / LoggedOutMenuDefaults — toggle File (Login vs Logoff, export, user maint)
     * and top-level Vessel / Reports / Specifications / Site Assignment visibility.
     */
    function syncMenuStripForSession() {
        var strip = document.getElementById("menuStrip");
        if (!strip) return;
        if (isSessionLoggedIn()) {
            strip.classList.remove("menu-strip-logged-out");
            strip.classList.add("menu-strip-logged-in");
        } else {
            strip.classList.remove("menu-strip-logged-in");
            strip.classList.add("menu-strip-logged-out");
        }
    }

    /**
     * frmMain.Login_Click — AutoLoginFlag = 1 → admin + Administrator; else frmLogin.ShowDialog.
     * On success: LoggedInMenuDefaults + SetupForm (refresh dashboard).
     */
    function performLoginFromMenu() {
        ensureSystemSettingsDefaults();
        var auto = state.systemSettings.autoLogin === true;
        if (auto) {
            state.currentUser = "Admin";
            var adm = state.users.filter(function (u) {
                return u.userId === "Admin" || u.name === "Admin";
            })[0];
            if (adm) {
                state.currentUser = adm.name;
                adm.lastLogon = new Date().toLocaleString();
            }
            syncMenuStripForSession();
            refreshUI({ staggerVesselReveal: true });
            toast("Logged in (auto login).");
            return;
        }
        closeMenus();
        var bdLoginEl = document.getElementById("backdropLogin");
        if (bdLoginEl) {
            bdLoginEl.classList.add("show");
            bdLoginEl.setAttribute("aria-hidden", "false");
        }
    }

    /**
     * frmMain.Logoff_Click — gbApplicationClosing; giAccessLevel ReadOnly; DisableVesselProcessingTimer;
     * LoggedOutMenuDefaults (clear gsLoggedInUserID, hide menus); clear vessel tabs / data tables.
     */
    function performLogoff() {
        state.applicationClosing = true;
        state.currentUser = null;
        state.tempGroupVesselIds = null;
        state.currentPage = 0;
        var grid = document.getElementById("vesselGrid");
        if (grid) grid.innerHTML = "";
        var tabStrip = document.getElementById("tabStrip");
        if (tabStrip) tabStrip.innerHTML = "";
        closeMenus();
        syncMenuStripForSession();
        saveState();
        updateTitleBar();
        updateMeasureMenuDisabled();
        state.applicationClosing = false;
        toast("Logged off.");
    }

    function computeAlarm(v) {
        return v.pctFull >= 90 ? "Pre-High Alarm" : "";
    }

    function applySimulatedMeasurementValues(v) {
        v.pctFull = Math.min(100, Math.max(5, v.pctFull + (Math.random() - 0.45) * 12));
        v.volumeCuFt = Math.round(800 + v.pctFull * 45);
        v.weightLb = Math.round(v.volumeCuFt * 42);
        v.heightFt = Math.round((4 + (v.pctFull / 100) * 10) * 100) / 100;
    }

    function measureVessel(v, opts) {
        opts = opts || {};
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to measure.");
            return;
        }
        if (state.vesselsReadOnly || !state.ebobServicesRunning) {
            toast(
                !state.ebobServicesRunning
                    ? "Database is read only — eBob services are not running."
                    : "Database is read only — close and restart Binventory to restore full access."
            );
            return;
        }
        var sid = parseInt(v.sensorTypeId, 10);
        if (isNaN(sid)) sid = 2;
        if (isSmartBobProtocolA(sid)) {
            runSmartBobProtocolAMeasurement(v, opts);
            return;
        }
        if (isSmartBobC100MB(sid)) {
            runSmartBobC100MBMeasurement(v, opts);
            return;
        }
        runGenericMeasurement(v, opts);
    }

    function measureAll() {
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to measure.");
            return;
        }
        if (state.vesselsReadOnly || !state.ebobServicesRunning) {
            toast(
                !state.ebobServicesRunning
                    ? "Database is read only — eBob services are not running."
                    : "Database is read only — close and restart Binventory to restore full access."
            );
            return;
        }
        var list = state.vessels.slice();
        if (list.length === 0) {
            toast("No vessels to measure.");
            return;
        }
        toast("Measuring all vessels…");
        var pending = list.length;
        function oneDone() {
            pending -= 1;
            if (pending <= 0) toast("All vessel measurements complete.");
        }
        list.forEach(function (v) {
            measureVessel(v, { silent: true, onDone: oneDone });
        });
    }

    /**
     * cardOpts.loading — Vessel.Designer.vb default lblStatusValue.Text = "Loading..." until bound (AddVesselsToTab).
     */
    function buildVesselCard(v, cardOpts) {
        cardOpts = cardOpts || {};
        var loading = !!cardOpts.loading;
        ensureVesselFields(v, 0);
        var pct = Math.round(v.pctFull);
        var alarm = computeAlarm(v) || "\u00A0";
        var headChecked = v.headroom ? " checked" : "";
        var m = vesselDisplayMetrics(v);
        var row = v.headroom ? m.headroom : m.product;
        var addrRaw = v.sensorAddress != null ? String(v.sensorAddress).trim() : "";
        var addrDisp = addrRaw !== "" ? addrRaw : "\u00A0";
        var dbRo = !!(state.vesselsReadOnly || !state.ebobServicesRunning);
        var statusText = dbRo ? "Database Read Only" : loading ? "Loading..." : v.status || "";
        var statusCls = dbRo ? " vessel-status-readonly" : loading ? " vessel-status-loading" : "";
        var measureDis = dbRo || loading ? " disabled" : "";
        var div = document.createElement("div");
        div.className = "vessel" + (loading ? " vessel--initializing" : "");
        div.dataset.vesselId = v.id;
        if (loading) div.setAttribute("aria-busy", "true");
        div.innerHTML =
            '<div class="vessel-head">' +
            '<div class="vessel-title-stack">' +
            '<div class="vessel-name-line" title="' +
            escapeHtml(v.product) +
            '">' +
            escapeHtml(v.name) +
            "</div>" +
            '<div class="vessel-sensor-line">' +
            escapeHtml(addrDisp) +
            "</div>" +
            "</div>" +
            '<span class="pic-multi" role="button" tabindex="0" title="Multi-vessel options" data-vessel-action="multi"></span>' +
            "</div>" +
            '<div class="vessel-inner">' +
            '<div class="vessel-col-left">' +
            '<div class="vessel-graphic-col">' +
            '<div class="silo" aria-hidden="true">' +
            '<div class="silo-sky"></div>' +
            '<div class="silo-fill" style="height:' +
            pct +
            '%"></div>' +
            "</div>" +
            '<div class="lbl-pct">' +
            pct +
            "% Full</div>" +
            "</div>" +
            "</div>" +
            '<div class="vessel-col-right">' +
            '<div class="kv">' +
            "<span>Height:</span><span>" +
            row.heightFt.toFixed(2) +
            " ft</span>" +
            "<span>Volume:</span><span>" +
            row.volumeCuFt.toLocaleString() +
            " cu ft</span>" +
            "<span>Weight:</span><span>" +
            row.weightLb.toLocaleString() +
            " lbs</span>" +
            "<span>Status:</span><span class=\"vessel-status-val" +
            statusCls +
            "\">" +
            escapeHtml(statusText) +
            "</span>" +
            "</div>" +
            '<label class="chk-row"><input type="checkbox" data-vessel-action="headroom"' +
            headChecked +
            '> Headroom Display</label>' +
            '<div class="lbl-alarm">' +
            escapeHtml(alarm) +
            "</div>" +
            '<div class="vessel-measure-row">' +
            '<button type="button" class="btn-measure"' +
            measureDis +
            ' data-vessel-action="measure">Measure</button>' +
            "</div>" +
            "</div>" +
            "</div>" +
            '<div class="lbl-last">' +
            escapeHtml(v.lastMeasurement) +
            "</div>";
        return div;
    }

    /**
     * Vessels shown on the dashboard grid (mirrors VesselUtility.ApplyTempGroup when gTempGroupVessels has rows).
     */
    function vesselsForDashboard() {
        var sid = state.currentWorkstationSiteId;
        var list = state.vessels.filter(function (v) {
            return v.siteId == null || v.siteId === sid;
        });
        list.sort(function (a, b) {
            var ao = a.sortOrder != null ? a.sortOrder : 0;
            var bo = b.sortOrder != null ? b.sortOrder : 0;
            return ao - bo;
        });
        var ids = state.tempGroupVesselIds;
        if (ids && ids.length > 0) {
            var byId = {};
            list.forEach(function (v) {
                byId[v.id] = v;
            });
            return ids
                .map(function (id) {
                    return byId[id];
                })
                .filter(function (v) {
                    return v;
                });
        }
        return list;
    }

    function pageCount() {
        return Math.max(1, Math.ceil(vesselsForDashboard().length / GRID_SIZE));
    }

    function renderTabs() {
        var strip = document.getElementById("tabStrip");
        strip.innerHTML = "";
        var n = pageCount();
        for (var p = 0; p < n; p++) {
            var b = document.createElement("button");
            b.type = "button";
            b.setAttribute("role", "tab");
            b.dataset.tab = String(p);
            b.textContent = "    Page " + (p + 1) + "    ";
            if (p === state.currentPage) {
                b.classList.add("active");
                b.setAttribute("aria-selected", "true");
            } else {
                b.setAttribute("aria-selected", "false");
            }
            b.addEventListener("click", function () {
                state.currentPage = parseInt(this.dataset.tab, 10);
                saveState();
                renderTabs();
                renderGrid();
            });
            strip.appendChild(b);
        }
    }

    /**
     * Each tile appears staggered with Status: Loading... (Vessel.Designer.vb), then binds to live status — softer than skeleton swap.
     */
    function renderGridStaggered(genSnapshot, opts) {
        opts = opts || {};
        var grid = document.getElementById("vesselGrid");
        grid.innerHTML = "";
        if (state.currentPage >= pageCount()) state.currentPage = 0;
        var start = state.currentPage * GRID_SIZE;
        var dash = vesselsForDashboard();
        var slice = dash.slice(start, start + GRID_SIZE);
        var appearDelay = opts.staggerDelayMs != null ? opts.staggerDelayMs : 44;
        var loadingDwell = opts.staggerLoadingDwellMs != null ? opts.staggerLoadingDwellMs : 420;
        var loadingCards = [];

        slice.forEach(function (v, idx) {
            window.setTimeout(function () {
                if (genSnapshot !== vesselGridRevealGen) return;
                var card = buildVesselCard(v, { loading: true });
                card.classList.add("vessel--appear");
                loadingCards[idx] = card;
                grid.appendChild(card);
            }, idx * appearDelay);
        });

        slice.forEach(function (v, idx) {
            window.setTimeout(function () {
                if (genSnapshot !== vesselGridRevealGen) return;
                var old = loadingCards[idx];
                if (!old || !old.parentNode) return;
                var card = buildVesselCard(v);
                card.classList.add("vessel--reveal-in");
                old.replaceWith(card);
            }, idx * appearDelay + loadingDwell);
        });
    }

    /** Time until last vessel tile finishes stagger reveal (matches renderGridStaggered). */
    function computeStaggerRevealDurationMs(opts) {
        opts = opts || {};
        var appearDelay = opts.staggerDelayMs != null ? opts.staggerDelayMs : 44;
        var loadingDwell = opts.staggerLoadingDwellMs != null ? opts.staggerLoadingDwellMs : 420;
        var start = state.currentPage * GRID_SIZE;
        var dash = vesselsForDashboard();
        var slice = dash.slice(start, start + GRID_SIZE);
        var n = slice.length;
        if (n === 0) return 400;
        /* After last stagger timeout, vessel--reveal-in runs ~0.48s (ebob.html) — keep overlay until that finishes */
        return (n - 1) * appearDelay + loadingDwell + 520;
    }

    function renderGrid(opts) {
        opts = opts || {};
        vesselGridRevealGen += 1;
        var genSnapshot = vesselGridRevealGen;
        if (opts.staggerVesselReveal) {
            renderGridStaggered(genSnapshot, opts);
            return;
        }
        var grid = document.getElementById("vesselGrid");
        grid.innerHTML = "";
        if (state.currentPage >= pageCount()) state.currentPage = 0;
        var start = state.currentPage * GRID_SIZE;
        var dash = vesselsForDashboard();
        var slice = dash.slice(start, start + GRID_SIZE);
        slice.forEach(function (v) {
            grid.appendChild(buildVesselCard(v));
        });
    }

    function updateMeasureMenuDisabled() {
        var btn = document.querySelector('button[data-action="measure-all"]');
        if (btn) {
            var dis = !isSessionLoggedIn() || state.vesselsReadOnly || !state.ebobServicesRunning;
            btn.disabled = dis;
            btn.classList.toggle("menu-btn-disabled", dis);
        }
    }

    /**
     * opts.staggerVesselReveal — cold start: each tile appears staggered with Status "Loading..." (Vessel.Designer.vb),
     * then settles to live status (opts.staggerDelayMs between tiles, opts.staggerLoadingDwellMs before bind).
     */
    function refreshUI(opts) {
        opts = opts || {};
        renderTabs();
        renderGrid(opts);
        updateTitleBar();
        refreshSiteAssignmentMenu();
        updateMeasureMenuDisabled();
        syncMenuStripForSession();
    }

    /**
     * engineRunning — eBob Engine Service up in services.msc sim.
     * When the engine stops, latch vesselsReadOnly (gbVesselsReadOnly in AppGlobals.vb / VesselUtility.vb);
     * starting services again does not clear it — same as production until Binventory is closed and restarted.
     */
    function setEbobServicesRunningFromSim(engineRunning) {
        var was = state.ebobServicesRunning;
        state.ebobServicesRunning = !!engineRunning;
        if (!state.ebobServicesRunning && was) {
            state.vesselsReadOnly = true;
        }
        refreshUI();
        if (!state.ebobServicesRunning && was) {
            toast(
                "Lost connection to database — read only mode. Close Binventory and restart the workstation after restoring eBob services."
            );
        } else if (state.ebobServicesRunning && !was) {
            if (state.vesselsReadOnly) {
                toast(
                    "eBob services are running. Close and restart Binventory to exit read-only mode."
                );
            } else {
                toast("Database connection restored — full access.");
            }
        }
    }

    /**
     * Mirrors frmMain.vb: LoadApplication → frmSplash.ShowDialog (Resources.Binventory_splash_screen_Feb_2023,
     * frmSplash.vb Timer 50 ticks + click) → login → tmrStartup 250 ms → SetupForm → ShowWait (frmWait.Show).
     * opts.includeSplash — default true (cold start). Set false to skip splash (e.g. testing).
     * opts.splashMs — auto-dismiss; default 5000 (~frmSplash 50 × 100 ms timer ticks).
     */
    function runBinventoryStartupSequence(onDone, opts) {
        opts = opts || {};
        var includeSplash = opts.includeSplash !== false;
        var waitMs = opts.waitMs != null ? opts.waitMs : 950;
        var splashMs = opts.splashMs != null ? opts.splashMs : 5000;

        var overlay = document.getElementById("startupOverlay");
        var splashEl = document.getElementById("startupFrmSplash");
        var waitEl = document.getElementById("startupFrmWait");
        var waitText = document.getElementById("startupWaitText");
        var pageWrap = document.getElementById("pageWrap");

        if (!overlay || !waitEl) {
            if (onDone) onDone();
            return;
        }

        if (pageWrap) pageWrap.classList.add("startup-sequence-active");

        function hideStartupPanels() {
            overlay.classList.remove("startup-overlay--visible");
            overlay.classList.remove("startup-overlay--splash-only");
            overlay.classList.remove("startup-overlay--wait-only");
            overlay.setAttribute("aria-hidden", "true");
            if (splashEl) splashEl.hidden = true;
            waitEl.hidden = true;
        }

        function finish() {
            hideStartupPanels();
            if (pageWrap) pageWrap.classList.remove("startup-sequence-active");
            if (onDone) onDone();
        }

        function runWaitPhase() {
            if (splashEl) splashEl.hidden = true;
            waitEl.hidden = false;
            overlay.classList.remove("startup-overlay--splash-only");
            overlay.classList.add("startup-overlay--wait-only");
            /* frmMain.vb SetupForm: ShowWait("Display is Loading. " & vbCrLf & "Please Wait...") */
            if (waitText) {
                waitText.textContent = "Display is Loading. \nPlease Wait...";
            }
            setTimeout(finish, waitMs);
        }

        overlay.setAttribute("aria-hidden", "false");

        if (includeSplash && splashEl) {
            splashEl.hidden = false;
            waitEl.hidden = true;
            overlay.classList.remove("startup-overlay--wait-only");
            overlay.classList.add("startup-overlay--splash-only");
            overlay.classList.add("startup-overlay--visible");
            var splashTimer = setTimeout(function () {
                splashEl.removeEventListener("click", onSplashClick);
                runWaitPhase();
            }, splashMs);
            function onSplashClick() {
                clearTimeout(splashTimer);
                splashEl.removeEventListener("click", onSplashClick);
                runWaitPhase();
            }
            splashEl.addEventListener("click", onSplashClick);
        } else {
            if (splashEl) splashEl.hidden = true;
            waitEl.hidden = true;
            overlay.classList.remove("startup-overlay--splash-only");
            overlay.classList.add("startup-overlay--wait-only");
            overlay.classList.add("startup-overlay--visible");
            setTimeout(runWaitPhase, 250);
        }
    }

    /**
     * Site Assignment (frmMain): Application.Restart — wait + staggered vessel refresh, no Binventory splash image.
     */
    function runSiteSwitchLoadingSequence(siteId, onDone) {
        if (!findSite(siteId)) {
            if (onDone) onDone();
            return;
        }

        /* Do not skip loading when prefers-reduced-motion is set (Windows "Animation effects" off sets this in Chromium).
         * Tile stagger may be suppressed by CSS @media (prefers-reduced-motion); overlay still shows. */

        var overlay = document.getElementById("startupOverlay");
        var splashEl = document.getElementById("startupFrmSplash");
        var waitEl = document.getElementById("startupFrmWait");
        var waitText = document.getElementById("startupWaitText");

        if (!overlay || !waitEl) {
            reloadWorkstationForSite(siteId);
            if (onDone) onDone();
            return;
        }

        if (splashEl) splashEl.hidden = true;
        waitEl.hidden = false;
        if (waitText) {
            waitText.textContent = "Display is Loading. \nPlease Wait...";
        }
        overlay.classList.remove("startup-overlay--splash-only");
        overlay.classList.add("startup-overlay--wait-only");
        overlay.classList.add("startup-overlay--site-switch");
        overlay.classList.add("startup-overlay--visible");
        overlay.setAttribute("aria-hidden", "false");
        /* Do NOT use startup-sequence-active here — it hides #vesselGrid (visibility:hidden), so stagger runs unseen. */

        function finishSiteSwitch() {
            overlay.classList.remove("startup-overlay--visible");
            overlay.classList.remove("startup-overlay--wait-only");
            overlay.classList.remove("startup-overlay--site-switch");
            overlay.setAttribute("aria-hidden", "true");
            waitEl.hidden = true;
            if (onDone) onDone();
        }

        function runReloadAndDismiss() {
            reloadWorkstationForSite(siteId, { skipRefreshUI: true });
            refreshUI({ staggerVesselReveal: true });
            var dwell = Math.max(computeStaggerRevealDurationMs({}), 1100);
            setTimeout(finishSiteSwitch, dwell);
        }

        requestAnimationFrame(function () {
            requestAnimationFrame(function () {
                setTimeout(runReloadAndDismiss, 280);
            });
        });
    }

    function closeNonAppBackdropsForExit() {
        var ids = [
            "backdropInfo",
            "backdropLogin",
            "backdropPrint",
            "backdropAbout",
            "backdropBinventoryMsg",
            "backdropAssignContacts",
            "backdropSnSetup",
            "backdropServicesMsc",
            "backdropUac",
            "backdropVsImportPaste"
        ];
        ids.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) {
                el.classList.remove("show");
                el.setAttribute("aria-hidden", "true");
            }
        });
        document.querySelectorAll(".backdrop-report-filter.show").forEach(function (el) {
            el.classList.remove("show");
        });
    }

    function exitToSimDesktop() {
        closeAppModal();
        closeNonAppBackdropsForExit();
        closeMenus();
        var pw = document.getElementById("pageWrap");
        if (pw) pw.classList.add("page-wrap--desktop-mode");
        var desk = document.getElementById("simDesktop");
        if (desk) {
            desk.classList.remove("sim-desktop--leaving");
            desk.classList.add("sim-desktop--entering");
            desk.hidden = false;
            desk.setAttribute("aria-hidden", "false");
            requestAnimationFrame(function () {
                requestAnimationFrame(function () {
                    desk.classList.remove("sim-desktop--entering");
                });
            });
        }
    }

    function launchEbobFromDesktop() {
        var pw = document.getElementById("pageWrap");
        var desk = document.getElementById("simDesktop");
        if (!pw || !desk) return;
        if (desk.hidden || desk.classList.contains("sim-desktop--leaving")) return;

        function startBinventoryAfterDesktop() {
            desk.classList.remove("sim-desktop--leaving");
            pw.classList.remove("page-wrap--desktop-mode");
            desk.hidden = true;
            desk.setAttribute("aria-hidden", "true");
            runBinventoryStartupSequence(function () {
                resetStateToFactoryDefaults();
                refreshUI({ staggerVesselReveal: true });
                toast("Connected to database — eBob Workstation is ready.");
            }, { includeSplash: true });
        }

        var reduceMotion =
            typeof window.matchMedia === "function" &&
            window.matchMedia("(prefers-reduced-motion: reduce)").matches;
        if (reduceMotion) {
            startBinventoryAfterDesktop();
            return;
        }

        desk.classList.add("sim-desktop--leaving");

        var finished = false;
        function finishDesktopFade() {
            if (finished) return;
            finished = true;
            desk.removeEventListener("transitionend", onTransitionEnd);
            if (fadeFallbackTimer != null) window.clearTimeout(fadeFallbackTimer);
            startBinventoryAfterDesktop();
        }

        function onTransitionEnd(ev) {
            if (ev.target !== desk) return;
            if (ev.propertyName !== "opacity" && ev.propertyName !== "transform") return;
            finishDesktopFade();
        }

        desk.addEventListener("transitionend", onTransitionEnd);
        var fadeFallbackTimer = window.setTimeout(finishDesktopFade, 480);
    }

    var bdApp = document.getElementById("backdropApp");
    var appModalShell = document.getElementById("appModalShell");
    var appModalTitle = document.getElementById("appModalTitle");
    var appModalBody = document.getElementById("appModalBody");
    var appModalFooter = document.getElementById("appModalFooter");

    /** Second layer (z-index) — frmScheduleMaintenance shown over frmEmailReports (WinForms ShowDialog stack). */
    var bdAppStack = document.getElementById("backdropAppStack");
    var appModalShellStack = document.getElementById("appModalShellStack");
    var appModalTitleStack = document.getElementById("appModalTitleStack");
    var appModalBodyStack = document.getElementById("appModalBodyStack");
    var appModalFooterStack = document.getElementById("appModalFooterStack");

    /** Third layer — Measurement Schedule Setup / Assign Groups over schedule list (parent stays visible). */
    var bdAppStack2 = document.getElementById("backdropAppStack2");
    var appModalShellStack2 = document.getElementById("appModalShellStack2");
    var appModalTitleStack2 = document.getElementById("appModalTitleStack2");
    var appModalBodyStack2 = document.getElementById("appModalBodyStack2");
    var appModalFooterStack2 = document.getElementById("appModalFooterStack2");

    function stripStackShellClasses(shell) {
        if (!shell) return;
        shell.classList.remove("modal-vessel-setup");
        shell.classList.remove("modal-sn-networks");
        shell.classList.remove("modal-site-maintenance");
        shell.classList.remove("modal-site-setup");
        shell.classList.remove("modal-system-setup");
        shell.classList.remove("modal-report-preview");
        shell.classList.remove("modal-report-crystal");
        shell.classList.remove("modal-scheduler-maint");
        shell.classList.remove("modal-measurement-schedule");
        shell.classList.remove("modal-temp-group");
        shell.classList.remove("modal-schedule-group-assign");
        shell.classList.remove("modal-vessel-group-assign");
        shell.classList.remove("modal-group-maint");
        shell.classList.remove("modal-user-maintenance");
        shell.classList.remove("modal-group-setup");
        shell.classList.remove("modal-email-setup");
        shell.classList.remove("modal-email-reports");
        shell.classList.remove("modal-win-toolwindow");
    }

    bdApp.addEventListener("click", function (e) {
        if (e.target.closest("[data-mock-report-print]")) {
            e.preventDefault();
            printMockReportWindow();
            return;
        }
        if (e.target.closest("[data-close-app]")) {
            if (appModalShell.classList.contains("modal-vessel-setup") && vmSetupEditingId) {
                closeVesselSetupDiscardChanges();
            } else {
                closeAppModal();
            }
        }
    });

    if (bdAppStack) {
        bdAppStack.addEventListener("click", function (e) {
            if (e.target.closest("[data-close-app-stack]")) {
                closeStackedAppModal();
            }
        });
    }

    if (bdAppStack2) {
        bdAppStack2.addEventListener("click", function (e) {
            if (e.target.closest("[data-close-app-stack2]")) {
                closeStackedAppModal2();
            }
        });
    }

    function updateStackedAppModalContent(title, bodyHtml, footerHtml) {
        if (!appModalTitleStack || !appModalBodyStack || !appModalFooterStack) return;
        if (title != null) appModalTitleStack.textContent = title;
        appModalBodyStack.innerHTML = bodyHtml;
        if (footerHtml === "") {
            appModalFooterStack.innerHTML = "";
            appModalShellStack.classList.add("modal-footer-hidden");
        } else {
            appModalShellStack.classList.remove("modal-footer-hidden");
            appModalFooterStack.innerHTML =
                footerHtml ||
                '<button type="button" class="secondary" data-close-app-stack>Close</button>';
        }
    }

    function updateStackedAppModalContent2(title, bodyHtml, footerHtml) {
        if (!appModalTitleStack2 || !appModalBodyStack2 || !appModalFooterStack2) return;
        if (title != null) appModalTitleStack2.textContent = title;
        appModalBodyStack2.innerHTML = bodyHtml;
        if (footerHtml === "") {
            appModalFooterStack2.innerHTML = "";
            appModalShellStack2.classList.add("modal-footer-hidden");
        } else {
            appModalShellStack2.classList.remove("modal-footer-hidden");
            appModalFooterStack2.innerHTML =
                footerHtml ||
                '<button type="button" class="secondary" data-close-app-stack2>Close</button>';
        }
    }

    function openStackedAppModal2(title, bodyHtml, footerHtml, modalClass) {
        if (!bdAppStack2 || !appModalShellStack2) return;
        updateStackedAppModalContent2(title, bodyHtml, footerHtml);
        stripStackShellClasses(appModalShellStack2);
        if (modalClass) {
            String(modalClass)
                .split(/\s+/)
                .forEach(function (c) {
                    if (c) appModalShellStack2.classList.add(c);
                });
        }
        bdAppStack2.classList.add("show");
        bdAppStack2.setAttribute("aria-hidden", "false");
    }

    function closeStackedAppModal2() {
        if (!bdAppStack2 || !appModalShellStack2) return;
        bdAppStack2.classList.remove("show");
        bdAppStack2.setAttribute("aria-hidden", "true");
        stripStackShellClasses(appModalShellStack2);
        appModalShellStack2.classList.remove("modal-footer-hidden");
    }

    /** Close MSS / Assign Groups layer only — schedule list stays on base or stack1. */
    function closeMssOrAssignLayer(emailFlowStack) {
        if (emailFlowStack) closeStackedAppModal2();
        else closeStackedAppModal();
    }

    function openStackedAppModal(title, bodyHtml, footerHtml, modalClass) {
        if (!bdAppStack || !appModalShellStack) return;
        closeStackedAppModal2();
        updateStackedAppModalContent(title, bodyHtml, footerHtml);
        stripStackShellClasses(appModalShellStack);
        if (modalClass) {
            String(modalClass)
                .split(/\s+/)
                .forEach(function (c) {
                    if (c) appModalShellStack.classList.add(c);
                });
        }
        bdAppStack.classList.add("show");
        bdAppStack.setAttribute("aria-hidden", "false");
    }

    function closeStackedAppModal() {
        closeStackedAppModal2();
        if (!bdAppStack || !appModalShellStack) return;
        bdAppStack.classList.remove("show");
        bdAppStack.setAttribute("aria-hidden", "true");
        stripStackShellClasses(appModalShellStack);
        appModalShellStack.classList.remove("modal-footer-hidden");
    }

    function openModalLayer(stacked, title, bodyHtml, footerHtml, modalClass) {
        if (stacked) openStackedAppModal(title, bodyHtml, footerHtml, modalClass);
        else openAppModal(title, bodyHtml, footerHtml, modalClass);
    }

    function closeModalLayer(stacked) {
        if (stacked) closeStackedAppModal();
        else closeAppModal();
    }

    function closeAppModal() {
        closeStackedAppModal();
        /* closeStackedAppModal clears stack2 then stack1 */
        dismissAssignContactsBackdrop();
        var vid = vmSetupEditingId;
        if (vid && vesselSetupSnapshotById[vid]) {
            var v = findVessel(vid);
            if (v) applyVesselSetupSnapshot(v, vesselSetupSnapshotById[vid]);
            delete vesselSetupSnapshotById[vid];
        }
        bdApp.classList.remove("show");
        appModalShell.classList.remove("modal-vessel-setup");
        appModalShell.classList.remove("modal-sn-networks");
        appModalShell.classList.remove("modal-site-maintenance");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-system-setup");
        appModalShell.classList.remove("modal-report-preview");
        appModalShell.classList.remove("modal-report-crystal");
        appModalShell.classList.remove("modal-scheduler-maint");
        appModalShell.classList.remove("modal-measurement-schedule");
        appModalShell.classList.remove("modal-temp-group");
        appModalShell.classList.remove("modal-schedule-group-assign");
        appModalShell.classList.remove("modal-vessel-group-assign");
        appModalShell.classList.remove("modal-group-maint");
        appModalShell.classList.remove("modal-user-maintenance");
        appModalShell.classList.remove("modal-group-setup");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-win-toolwindow");
        vmSetupEditingId = null;
    }

    function updateAppModalContent(title, bodyHtml, footerHtml) {
        if (title != null) appModalTitle.textContent = title;
        appModalBody.innerHTML = bodyHtml;
        if (footerHtml === "") {
            appModalFooter.innerHTML = "";
            appModalShell.classList.add("modal-footer-hidden");
        } else {
            appModalShell.classList.remove("modal-footer-hidden");
            appModalFooter.innerHTML = footerHtml || '<button type="button" class="secondary" data-close-app>Close</button>';
        }
    }

    function openAppModal(title, bodyHtml, footerHtml, modalClass) {
        closeStackedAppModal(); /* stack2 + stack1 */
        updateAppModalContent(title, bodyHtml, footerHtml);
        appModalShell.classList.remove("modal-vessel-setup");
        appModalShell.classList.remove("modal-sn-networks");
        appModalShell.classList.remove("modal-site-maintenance");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-system-setup");
        appModalShell.classList.remove("modal-report-preview");
        appModalShell.classList.remove("modal-report-crystal");
        appModalShell.classList.remove("modal-scheduler-maint");
        appModalShell.classList.remove("modal-measurement-schedule");
        appModalShell.classList.remove("modal-temp-group");
        appModalShell.classList.remove("modal-schedule-group-assign");
        appModalShell.classList.remove("modal-vessel-group-assign");
        appModalShell.classList.remove("modal-group-maint");
        appModalShell.classList.remove("modal-user-maintenance");
        appModalShell.classList.remove("modal-group-setup");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-win-toolwindow");
        if (modalClass) {
            String(modalClass)
                .split(/\s+/)
                .forEach(function (c) {
                    if (c) appModalShell.classList.add(c);
                });
        }
        bdApp.classList.add("show");
    }

    function printMockReportWindow() {
        var title = appModalTitle ? appModalTitle.textContent : "Report";
        var body = appModalBody ? appModalBody.innerHTML : "";
        var w = window.open("", "_blank");
        if (!w) {
            toast("Pop-up blocked — allow pop-ups to print this report.");
            return;
        }
        var crystal = body.indexOf("rpt-crystal-shell") !== -1;
        var sty =
            "<style>body{font-family:Arial,Helvetica,Tahoma,sans-serif;font-size:11px;padding:12px;color:#000;margin:0;background:#fff;}" +
            "h1{font-size:15px;margin:0 0 12px;}" +
            ".rpt-mock-meta,.rpt-mock-meta strong{font-size:12px;margin:0 0 8px;}.rpt-mock-note{color:#666;font-size:11px;margin:0 0 12px;}" +
            ".rpt-mock-bar-row{display:flex;align-items:center;gap:8px;margin:8px 0;}.rpt-mock-bar-track{flex:1;height:14px;background:#e8e8e8;border:1px solid #ccc;}.rpt-mock-bar-fill{height:100%;background:#3d6cb0;}.rpt-mock-bar-name{min-width:140px;font-size:11px;}.rpt-mock-chart-svg{width:100%;max-width:560px;display:block;}" +
            ".rpt-crystal-shell{font-family:Arial,Helvetica,Tahoma,sans-serif;font-size:11px;color:#000;}" +
            ".rpt-crystal-toolbar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;background:linear-gradient(180deg,#f5f5f5 0%,#e4e4e4 100%);border:1px solid #b0b0b0;border-bottom:none;padding:4px 8px;}" +
            ".rpt-crystal-toolbar-btns{display:flex;flex-wrap:wrap;align-items:center;gap:2px;font-size:11px;}" +
            ".rpt-crystal-tb-btn{display:inline-block;padding:2px 8px;border:1px solid #b8b8b8;background:#fafafa;cursor:default;border-radius:1px;}" +
            ".rpt-crystal-tb-sep{color:#999;padding:0 4px;}" +
            ".rpt-crystal-toolbar-brand{font-size:10px;color:#666;letter-spacing:0.02em;}" +
            ".rpt-crystal-tabrow{background:#ececec;border:1px solid #b0b0b0;border-top:none;padding:2px 8px 0;}" +
            ".rpt-crystal-tab-active{display:inline-block;padding:4px 12px;background:#fff;border:1px solid #b0b0b0;border-bottom:none;font-weight:600;}" +
            ".rpt-crystal-canvas{background:#c8c8c8;padding:16px 20px 24px;border:1px solid #b0b0b0;border-top:none;}" +
            ".rpt-crystal-page{background:#fff;box-shadow:0 1px 4px rgba(0,0,0,0.25);padding:16px 20px 20px;max-width:100%;}" +
            ".rpt-crystal-header-grid{display:grid;grid-template-columns:1fr 1.4fr 1fr;gap:8px;margin-bottom:10px;font-size:11px;align-items:start;}" +
            ".rpt-crystal-h-left{text-align:left;line-height:1.45;}" +
            ".rpt-crystal-h-center{text-align:center;line-height:1.35;}" +
            ".rpt-crystal-h-center strong{font-size:12px;}" +
            ".rpt-crystal-h-right{text-align:right;line-height:1.35;}" +
            ".rpt-crystal-header-line{border:none;border-top:1px solid #000;margin:0 0 8px;}" +
            ".rpt-crystal-data-table{width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed;}" +
            ".rpt-crystal-data-table th{text-align:left;font-weight:600;padding:4px 6px;border-bottom:1px solid #000;background:transparent;}" +
            ".rpt-crystal-data-table td{padding:6px 6px 8px;vertical-align:top;border-bottom:1px solid #ccc;word-wrap:break-word;}" +
            ".rpt-crystal-desc-cell{width:28%;}" +
            ".rpt-crystal-stack{margin-top:6px;font-size:10px;line-height:1.35;color:#222;}" +
            ".rpt-crystal-table-end{border:none;border-top:2px solid #000;margin:12px 0 0;}" +
            ".rpt-crystal-statusbar{background:#ececec;border:1px solid #b0b0b0;border-top:none;padding:4px 10px;font-size:11px;color:#000;}" +
            ".rpt-mock-table{width:100%;border-collapse:collapse;font-size:12px;}" +
            ".rpt-mock-table th,.rpt-mock-table td{border:1px solid #c0c0c0;padding:6px 8px;}" +
            ".rpt-mock-table th{background:#f0f0f0;}</style>";
        w.document.write(
            "<!DOCTYPE html><html><head><meta charset='utf-8'><title>" +
                escapeHtml(title) +
                "</title>" +
                sty +
                "</head><body" +
                (crystal ? ' style="background:#fff;padding:0"' : "") +
                ">"
        );
        if (!crystal) {
            w.document.write("<h1>" + escapeHtml(title) + "</h1>");
        }
        w.document.write(body);
        w.document.write("</body></html>");
        w.document.close();
        w.focus();
        setTimeout(function () {
            w.print();
        }, 200);
    }

    function downloadTextFile(filename, text, mime) {
        var blob = new Blob([text], { type: mime || "text/csv;charset=utf-8" });
        var a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function findVessel(id) {
        for (var i = 0; i < state.vessels.length; i++) {
            if (state.vessels[i].id === id) return state.vessels[i];
        }
        return null;
    }

    function findContactById(id) {
        var c;
        for (c = 0; c < state.contacts.length; c++) {
            if (state.contacts[c].id === id) return state.contacts[c];
        }
        return null;
    }

    function contactNameParts(c) {
        var parts = String(c.name || "").trim().split(/\s+/);
        return { first: parts[0] || "", last: parts.length > 1 ? parts.slice(1).join(" ") : "" };
    }

    function ensureVesselContactIds(v) {
        if (!v) return;
        if (!Array.isArray(v.vesselContactIds)) v.vesselContactIds = [];
    }

    function sortContactIdsStable(ids) {
        var order = {};
        state.contacts.forEach(function (c, i) {
            order[c.id] = i;
        });
        return ids.slice().sort(function (a, b) {
            return (order[a] != null ? order[a] : 999) - (order[b] != null ? order[b] : 999);
        });
    }

    function contactRowHtml(c) {
        var nm = contactNameParts(c);
        return (
            '<tr data-cid="' +
            escapeHtml(c.id) +
            '"><td>' +
            escapeHtml(nm.first) +
            "</td><td>" +
            escapeHtml(nm.last) +
            "</td><td>" +
            escapeHtml(c.email || "") +
            "</td></tr>"
        );
    }

    function buildVesselAssignedContactRowsHtml(v) {
        ensureVesselContactIds(v);
        var rows = [];
        v.vesselContactIds.forEach(function (cid) {
            var c = findContactById(cid);
            if (c) rows.push(contactRowHtml(c));
        });
        return rows.length
            ? rows.join("")
            : '<tr><td colspan="3" class="vs-muted">No contacts assigned.</td></tr>';
    }

    var assignContactsState = null;

    function renderAssignContactsTables() {
        if (!assignContactsState) return;
        var un = document.getElementById("ac_tbody_un");
        var as = document.getElementById("ac_tbody_as");
        if (!un || !as) return;
        un.innerHTML = assignContactsState.unassigned
            .map(function (id) {
                var c = findContactById(id);
                return c ? contactRowHtml(c) : "";
            })
            .filter(Boolean)
            .join("");
        if (!un.innerHTML.trim()) {
            un.innerHTML = '<tr><td colspan="3" class="vs-muted ac-grid-empty">&nbsp;</td></tr>';
        }
        as.innerHTML = assignContactsState.assigned
            .map(function (id) {
                var c = findContactById(id);
                return c ? contactRowHtml(c) : "";
            })
            .filter(Boolean)
            .join("");
        if (!as.innerHTML.trim()) {
            as.innerHTML = '<tr><td colspan="3" class="vs-muted ac-grid-empty">&nbsp;</td></tr>';
        }
    }

    function syncAssignListsFromAssigned() {
        if (!assignContactsState) return;
        var allIds = state.contacts.map(function (c) {
            return c.id;
        });
        var as = sortContactIdsStable(
            assignContactsState.assigned.filter(function (id) {
                return allIds.indexOf(id) >= 0;
            })
        );
        assignContactsState.assigned = as;
        assignContactsState.unassigned = sortContactIdsStable(
            allIds.filter(function (id) {
                return as.indexOf(id) < 0;
            })
        );
    }

    function openAssignContactsDialog() {
        var v = findVessel(vmSetupEditingId);
        if (!v) return;
        ensureVesselContactIds(v);
        var allIds = state.contacts.map(function (c) {
            return c.id;
        });
        var assigned = sortContactIdsStable(v.vesselContactIds.slice());
        var unassigned = sortContactIdsStable(
            allIds.filter(function (id) {
                return assigned.indexOf(id) < 0;
            })
        );
        assignContactsState = { vesselId: v.id, unassigned: unassigned, assigned: assigned };
        renderAssignContactsTables();
        var bac = document.getElementById("backdropAssignContacts");
        if (bac) bac.classList.add("show");
    }

    function closeAssignContactsDialog() {
        var v = assignContactsState && findVessel(assignContactsState.vesselId);
        if (v) {
            ensureVesselContactIds(v);
            v.vesselContactIds = sortContactIdsStable(assignContactsState.assigned.slice());
        }
        assignContactsState = null;
        var bac = document.getElementById("backdropAssignContacts");
        if (bac) bac.classList.remove("show");
        refreshVesselEmailContactsTable();
    }

    function dismissAssignContactsBackdrop() {
        assignContactsState = null;
        var bac = document.getElementById("backdropAssignContacts");
        if (bac) bac.classList.remove("show");
    }

    function assignContactsSelectedCid(tbody) {
        if (!tbody) return null;
        var tr = tbody.querySelector("tr.ac-sel[data-cid]");
        return tr ? tr.getAttribute("data-cid") : null;
    }

    function assignContactsAddIds(ids) {
        if (!assignContactsState) return;
        ids.forEach(function (id) {
            if (assignContactsState.assigned.indexOf(id) < 0) assignContactsState.assigned.push(id);
        });
        syncAssignListsFromAssigned();
        renderAssignContactsTables();
    }

    function assignContactsRemoveIds(ids) {
        if (!assignContactsState) return;
        assignContactsState.assigned = assignContactsState.assigned.filter(function (id) {
            return ids.indexOf(id) < 0;
        });
        syncAssignListsFromAssigned();
        renderAssignContactsTables();
    }

    function wireAssignContactsDialog() {
        var bd = document.getElementById("backdropAssignContacts");
        if (!bd || bd._acWired) return;
        bd._acWired = true;
        var wr = document.getElementById("ac_assign_wrap");
        if (wr) {
            wr.addEventListener("click", function (e) {
                var tr = e.target.closest("tr[data-cid]");
                if (!tr) return;
                var tb = tr.parentElement;
                if (tb.id !== "ac_tbody_un" && tb.id !== "ac_tbody_as") return;
                tb.querySelectorAll("tr[data-cid]").forEach(function (x) {
                    x.classList.remove("ac-sel");
                });
                tr.classList.add("ac-sel");
            });
        }
        document.getElementById("ac_btn_add").addEventListener("click", function () {
            var cid = assignContactsSelectedCid(document.getElementById("ac_tbody_un"));
            if (cid) assignContactsAddIds([cid]);
        });
        document.getElementById("ac_btn_remove").addEventListener("click", function () {
            var cid = assignContactsSelectedCid(document.getElementById("ac_tbody_as"));
            if (cid) assignContactsRemoveIds([cid]);
        });
        document.getElementById("ac_btn_add_all").addEventListener("click", function () {
            if (!assignContactsState) return;
            assignContactsState.assigned = sortContactIdsStable(
                state.contacts.map(function (c) {
                    return c.id;
                })
            );
            syncAssignListsFromAssigned();
            renderAssignContactsTables();
        });
        document.getElementById("ac_btn_remove_all").addEventListener("click", function () {
            if (!assignContactsState) return;
            assignContactsState.assigned = [];
            syncAssignListsFromAssigned();
            renderAssignContactsTables();
        });
        document.getElementById("ac_btn_close").addEventListener("click", function () {
            closeAssignContactsDialog();
        });
    }

    function refreshVesselEmailContactsTable() {
        var tbody = document.getElementById("vs_email_contacts_tbody");
        var v = findVessel(vmSetupEditingId);
        if (!tbody || !v) return;
        tbody.innerHTML = buildVesselAssignedContactRowsHtml(v);
    }

    var vmSelectedId = null;
    var vmSetupEditingId = null;
    /** Deep copy of vessel state when Vessel Setup opens — restored on Cancel / discard so edits don't stick without Save. */
    var vesselSetupSnapshotById = {};

    function takeVesselSetupSnapshot(v) {
        return JSON.parse(JSON.stringify(v));
    }

    function applyVesselSetupSnapshot(v, snap) {
        if (!v || !snap) return;
        var plain = JSON.parse(JSON.stringify(snap));
        Object.keys(plain).forEach(function (key) {
            v[key] = plain[key];
        });
    }

    function closeVesselSetupDiscardChanges() {
        var vid = vmSetupEditingId;
        if (vid && vesselSetupSnapshotById[vid]) {
            var v = findVessel(vid);
            if (v) applyVesselSetupSnapshot(v, vesselSetupSnapshotById[vid]);
            delete vesselSetupSnapshotById[vid];
        }
        closeVesselSetupToMaintenance();
    }

    function vesselsSortedList() {
        return state.vessels.slice().sort(function (a, b) {
            return (a.sortOrder || 0) - (b.sortOrder || 0);
        });
    }

    function maxSortOrder() {
        var m = 0;
        state.vessels.forEach(function (v) {
            if (v.sortOrder != null && v.sortOrder > m) m = v.sortOrder;
        });
        return m;
    }

    function nextVesselNumericId() {
        var max = 0;
        state.vessels.forEach(function (v) {
            ensureVesselFields(v, 0);
            if (v.vesselNumericId > max) max = v.vesselNumericId;
        });
        return max + 1;
    }

    function sensorNetworkDisplayName(networkId) {
        var n = null;
        for (var i = 0; i < state.sensorNetworks.length; i++) {
            if (state.sensorNetworks[i].id === networkId) {
                n = state.sensorNetworks[i];
                break;
            }
        }
        return n ? n.name : "";
    }

    function removeVesselFromGroupsAndSchedules(vid) {
        state.groups.forEach(function (g) {
            g.vesselIds = g.vesselIds.filter(function (id) { return id !== vid; });
        });
        state.schedules.forEach(function (s) {
            s.vesselIds = s.vesselIds.filter(function (id) { return id !== vid; });
        });
    }

    function vmRefreshToolbar() {
        var hasRows = state.vessels.length > 0;
        var sel = vmSelectedId ? findVessel(vmSelectedId) : null;
        var rowSelected = !!sel;
        var moreThanOne = state.vessels.length > 1;
        var sorted = vesselsSortedList();
        var idx = rowSelected
            ? sorted.findIndex(function (v) {
                return v.id === vmSelectedId;
            })
            : -1;

        function set(id, disabled) {
            var el = document.getElementById(id);
            if (el) el.disabled = disabled;
        }

        set("vmBtnSelect", !rowSelected);
        set("vmBtnAddNew", false);
        set("vmBtnAddUsing", !rowSelected);
        set("vmBtnDelete", !rowSelected);
        set("vmBtnSortName", !moreThanOne);
        set("vmBtnSortContents", !moreThanOne);
        set("vmBtnMoveUp", !rowSelected || !moreThanOne || idx <= 0);
        set("vmBtnMoveDown", !rowSelected || !moreThanOne || idx < 0 || idx >= sorted.length - 1);
    }

    function buildVmListHtml() {
        var sorted = vesselsSortedList();
        var rows = sorted
            .map(function (v) {
                var sel = v.id === vmSelectedId ? " selected" : "";
                var net = sensorNetworkDisplayName(v.sensorNetworkId);
                var contents = v.contents != null ? v.contents : v.product;
                return (
                    "<tr class=\"vm-row" + sel + "\" data-vid=\"" + escapeHtml(v.id) + "\">" +
                    "<td>" + escapeHtml(v.name) + "</td>" +
                    "<td>" + escapeHtml(String(contents).substring(0, 120)) + "</td>" +
                    "<td class=\"vm-col-network\">" + escapeHtml(net) + "</td>" +
                    "</tr>"
                );
            })
            .join("");
        return (
            '<div class="vm-shell">' +
            '<div class="vm-header-title">Vessel Maintenance</div>' +
            '<div class="vm-main-row">' +
            '<div class="vm-table-panel">' +
            '<table class="data-table vm-grid-table" id="vmTable">' +
            "<thead><tr><th>Vessel Name</th><th>Contents</th><th>Sensor Network</th></tr></thead>" +
            "<tbody>" +
            (rows || '<tr><td colspan="3" class="muted">No vessels.</td></tr>') +
            "</tbody></table>" +
            "</div>" +
            '<div class="vm-actions-col">' +
            '<button type="button" class="vm-action-btn" id="vmBtnSelect">Select</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnAddNew">Add New</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnAddUsing">Add New Using </button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnDelete">Delete</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnSortName">Sort by Name</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnSortContents">Sort by Contents</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnMoveUp">Move Up</button>' +
            '<button type="button" class="vm-action-btn" id="vmBtnMoveDown">Move Down</button>' +
            '<button type="button" class="vm-action-btn vm-close-btn" id="vmBtnClose" data-close-app>Close</button>' +
            "</div>" +
            "</div>" +
            "</div>"
        );
    }

    function bindVmList() {
        var tbody = document.querySelector("#vmTable tbody");
        if (!tbody) return;

        tbody.querySelectorAll("tr.vm-row").forEach(function (tr) {
            tr.addEventListener("click", function () {
                vmSelectedId = tr.dataset.vid;
                tbody.querySelectorAll("tr").forEach(function (r) { r.classList.remove("selected"); });
                tr.classList.add("selected");
                vmRefreshToolbar();
            });
            tr.addEventListener("dblclick", function () {
                vmSelectedId = tr.dataset.vid;
                openVmSetupView(vmSelectedId);
            });
        });

        document.getElementById("vmBtnSelect").addEventListener("click", function () {
            if (vmSelectedId) openVmSetupView(vmSelectedId);
        });
        document.getElementById("vmBtnAddNew").addEventListener("click", function () {
            var nv = seedVessel(state.vessels.length);
            nv.id = uid("v");
            nv.vesselNumericId = nextVesselNumericId();
            nv.sortOrder = maxSortOrder() + 1;
            nv.name = "New vessel";
            nv.sensorNetworkId = state.sensorNetworks[0] ? state.sensorNetworks[0].id : null;
            ensureVesselFields(nv, state.vessels.length);
            state.vessels.push(nv);
            vmSelectedId = nv.id;
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
            refreshUI();
            openVmSetupView(nv.id);
        });
        document.getElementById("vmBtnAddUsing").addEventListener("click", function () {
            if (!vmSelectedId) return;
            var base = findVessel(vmSelectedId);
            if (!base) return;
            var nv = JSON.parse(JSON.stringify(base));
            nv.id = uid("v");
            nv.vesselNumericId = nextVesselNumericId();
            nv.sortOrder = maxSortOrder() + 1;
            nv.name = (base.name || "Vessel") + " (copy)";
            ensureVesselFields(nv, state.vessels.length);
            assignUniqueSensorAddressForVessel(nv);
            state.vessels.push(nv);
            vmSelectedId = nv.id;
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
            refreshUI();
            openVmSetupView(nv.id);
        });
        document.getElementById("vmBtnDelete").addEventListener("click", function () {
            if (!vmSelectedId || state.vessels.length <= 1) {
                toast("At least one vessel must remain.");
                return;
            }
            var v = findVessel(vmSelectedId);
            var msg =
                "Are you sure you want to delete the vessel named '" +
                (v ? String(v.name) : "") +
                "'?\n\n" +
                "IMPORTANT:  Deleting this vessel will remove it from any groups and schedules and also delete it's measurement history.";
            if (!confirm(msg)) return;
            removeVesselFromGroupsAndSchedules(vmSelectedId);
            state.vessels = state.vessels.filter(function (x) { return x.id !== vmSelectedId; });
            vmSelectedId = state.vessels[0] ? state.vessels[0].id : null;
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
            refreshUI();
            toast("Vessel deleted (simulation).");
        });
        document.getElementById("vmBtnSortName").addEventListener("click", function () {
            if (state.vessels.length < 2) return;
            state.vessels.sort(function (a, b) {
                var an = (a.name || "").toLowerCase();
                var bn = (b.name || "").toLowerCase();
                if (an !== bn) return an < bn ? -1 : 1;
                var ac = (a.contents || a.product || "").toLowerCase();
                var bc = (b.contents || b.product || "").toLowerCase();
                return ac < bc ? -1 : ac > bc ? 1 : 0;
            });
            state.vessels.forEach(function (v, i) {
                v.sortOrder = i + 1;
            });
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
        });
        document.getElementById("vmBtnSortContents").addEventListener("click", function () {
            if (state.vessels.length < 2) return;
            state.vessels.sort(function (a, b) {
                var ac = (a.contents || a.product || "").toLowerCase();
                var bc = (b.contents || b.product || "").toLowerCase();
                if (ac !== bc) return ac < bc ? -1 : 1;
                var an = (a.name || "").toLowerCase();
                var bn = (b.name || "").toLowerCase();
                return an < bn ? -1 : an > bn ? 1 : 0;
            });
            state.vessels.forEach(function (v, i) {
                v.sortOrder = i + 1;
            });
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
        });
        document.getElementById("vmBtnMoveUp").addEventListener("click", function () {
            if (!vmSelectedId) return;
            var sorted = vesselsSortedList();
            var idx = sorted.findIndex(function (v) { return v.id === vmSelectedId; });
            if (idx <= 0) return;
            var o = sorted[idx].sortOrder;
            sorted[idx].sortOrder = sorted[idx - 1].sortOrder;
            sorted[idx - 1].sortOrder = o;
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
        });
        document.getElementById("vmBtnMoveDown").addEventListener("click", function () {
            if (!vmSelectedId) return;
            var sorted = vesselsSortedList();
            var idx = sorted.findIndex(function (v) { return v.id === vmSelectedId; });
            if (idx < 0 || idx >= sorted.length - 1) return;
            var o = sorted[idx].sortOrder;
            sorted[idx].sortOrder = sorted[idx + 1].sortOrder;
            sorted[idx + 1].sortOrder = o;
            saveState();
            updateAppModalContent(
                "Vessel Maintenance — Binventory Workstation",
                buildVmListHtml(),
                ""
            );
            bindVmList();
            vmRefreshToolbar();
        });

        vmRefreshToolbar();
    }

    function buildVmSetupProductOptions(current) {
        var cur = current || "";
        var seen = {};
        var parts = PRODUCTS.map(function (p) {
            seen[p] = true;
            return "<option" + (p === cur ? " selected" : "") + ">" + escapeHtml(p) + "</option>";
        });
        if (cur && !seen[cur]) {
            parts.unshift("<option selected>" + escapeHtml(cur) + "</option>");
        }
        return parts.join("");
    }

    function getSensorNetworkById(id) {
        var arr = state.sensorNetworks || [];
        for (var i = 0; i < arr.length; i++) {
            if (arr[i].id === id) return arr[i];
        }
        return null;
    }

    /** Display like eBob: "Name (Protocol on COM1)" or IP for remote. */
    function formatNetworkSelectLabel(n) {
        if (!n) return "";
        if (!n.name) n.name = "Sensor Network";
        if (!n.protocol) n.protocol = "Protocol A";
        if (!n.interface) n.interface = "COM1";
        var iface =
            n.interfaceMode === "remote" && (n.remoteIp || "").toString().trim()
                ? String(n.remoteIp).trim() +
                  (n.remotePort ? ":" + String(n.remotePort).trim() : "")
                : n.interface;
        return n.name + " (" + n.protocol + " on " + iface + ")";
    }

    /** Modbus/RTU dropdown order — matches Binventory workstation (SortOrder / tblDeviceType). */
    var MODBUS_RTU_SENSOR_TYPE_ORDER = [13, 4, 5, 6, 7, 8, 9, 10, 11, 12];

    function deviceTypeById(id) {
        var tid = parseInt(id, 10);
        if (isNaN(tid)) return null;
        for (var i = 0; i < SENSOR_DEVICE_TYPES.length; i++) {
            if (SENSOR_DEVICE_TYPES[i].id === tid) return SENSOR_DEVICE_TYPES[i];
        }
        return null;
    }

    /**
     * Map sensor network protocol to allowed device type IDs (tutorial mock).
     * Modbus/RTU uses a fixed list order for the same UX as the desktop app.
     */
    function sensorTypeIdsForNetwork(network) {
        if (!network) {
            return SENSOR_DEVICE_TYPES.map(function (x) {
                return x.id;
            });
        }
        var p = network.protocol || "Protocol A";
        if (p === "Protocol A") {
            return [1, 3];
        }
        if (p === "Protocol B") {
            return [2];
        }
        if (p === "Modbus/RTU") {
            return MODBUS_RTU_SENSOR_TYPE_ORDER.slice();
        }
        if (p === "HART Protocol") {
            return [16];
        }
        if (p === "SPL-100 Push") {
            return [14];
        }
        if (p === "SPL-200 Push") {
            return [15];
        }
        return SENSOR_DEVICE_TYPES.map(function (x) {
            return x.id;
        });
    }

    function buildSensorTypeOptionsForNetwork(networkId, selectedTypeId) {
        var n = getSensorNetworkById(networkId);
        var allowed = sensorTypeIdsForNetwork(n);
        var seen = {};
        allowed.forEach(function (id) {
            seen[id] = true;
        });
        var sel = selectedTypeId != null ? parseInt(selectedTypeId, 10) : NaN;
        if (isNaN(sel) || !seen[sel]) {
            sel = allowed[0] != null ? allowed[0] : 1;
        }
        var parts = [];
        allowed.forEach(function (id) {
            var row = deviceTypeById(id);
            if (!row) return;
            parts.push(
                "<option value=\"" +
                    escapeHtml(String(row.id)) +
                    "\"" +
                    (row.id === sel ? " selected" : "") +
                    ">" +
                    escapeHtml(row.name) +
                    "</option>"
            );
        });
        return { html: parts.join(""), selectedId: sel };
    }

    function buildVmSetupNetworkOptions(selectedId) {
        var parts = [];
        (state.sensorNetworks || []).forEach(function (n) {
            parts.push(
                "<option value=\"" +
                    escapeHtml(n.id) +
                    "\"" +
                    (n.id === selectedId ? " selected" : "") +
                    ">" +
                    escapeHtml(formatNetworkSelectLabel(n)) +
                    "</option>"
            );
        });
        if (!parts.length) {
            parts.push('<option value="">—</option>');
        }
        return parts.join("");
    }

    function buildSelectOptionsStringList(items, selected) {
        return items
            .map(function (item) {
                var sel = item === selected ? " selected" : "";
                return "<option" + sel + ">" + escapeHtml(item) + "</option>";
            })
            .join("");
    }

    function buildKeyedSelectOptions(list, selectedId, valueKey, labelKey) {
        return list
            .map(function (row) {
                var id = row[valueKey];
                var sel = String(id) === String(selectedId) ? " selected" : "";
                return (
                    "<option value=\"" +
                    escapeHtml(String(id)) +
                    "\"" +
                    sel +
                    ">" +
                    escapeHtml(row[labelKey]) +
                    "</option>"
                );
            })
            .join("");
    }

    function buildFillColorOptions(selected) {
        return FILL_COLOR_NAMES.map(function (name) {
            return (
                "<option" +
                (name === selected ? " selected" : "") +
                ">" +
                escapeHtml(name) +
                "</option>"
            );
        }).join("");
    }

    function buildVesselSetupBody(v) {
        var nm = v.name || "";
        var hLabel = heightUnitsLabel();
        var contents = v.contents != null ? v.contents : v.product;
        var fillSel = v.fillColor || "Light Blue";
        var preview = fillColorPreviewCss(fillSel);

        var tabStrip =
            '<div class="vs-tabstrip" role="tablist">' +
            '<button type="button" class="vs-tab active" role="tab" data-vs-tab="display" aria-selected="true">Display</button>' +
            '<button type="button" class="vs-tab" role="tab" data-vs-tab="typeshape" aria-selected="false">Type/Shape</button>' +
            '<button type="button" class="vs-tab" role="tab" data-vs-tab="contents" aria-selected="false">Contents</button>' +
            '<button type="button" class="vs-tab" role="tab" data-vs-tab="sensor" aria-selected="false">Sensor</button>' +
            '<button type="button" class="vs-tab" role="tab" data-vs-tab="alarms" aria-selected="false">Alarms</button>' +
            '<button type="button" class="vs-tab" role="tab" data-vs-tab="email" aria-selected="false">Email Notifications</button>' +
            "</div>";

        var tabDisplay =
            '<div class="vs-tab-panel active" data-vs-panel="display" role="tabpanel">' +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Vessel Name and Fill Color</legend>' +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_name">Vessel Name:</label>' +
            '<div class="vs-field-control">' +
            '<input type="text" id="vs_name" class="vs-input" maxlength="10" value="' +
            escapeHtml(nm) +
            '">' +
            "</div></div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_color">Fill Color:</label>' +
            '<div class="vs-field-control vs-color-row">' +
            '<select id="vs_color" class="vs-select">' +
            buildFillColorOptions(fillSel) +
            "</select>" +
            '<div id="vs_color_preview" class="vs-color-preview" style="background:' +
            preview +
            '"></div>' +
            "</div></div>" +
            "</fieldset>" +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Displayed Units of Measure</legend>' +
            '<div class="vs-field-row">' +
            '<span class="vs-lbl">Height Units:</span>' +
            '<div class="vs-field-control"><span id="vs_height_units" class="vs-static-units">' +
            escapeHtml(hLabel) +
            "</span></div></div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_vol_units">Volume Units:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_vol_units" class="vs-select">' +
            buildSelectOptionsStringList(VOLUME_DISPLAY_UNITS, v.volumeDisplayUnits) +
            "</select></div></div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_wt_units">Weight Units:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_wt_units" class="vs-select">' +
            buildSelectOptionsStringList(WEIGHT_DISPLAY_UNITS, v.weightDisplayUnits) +
            "</select></div></div>" +
            "</fieldset>" +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Options</legend>' +
            '<label class="vs-check">' +
            '<input type="checkbox" id="vs_headroom"' +
            (v.defaultHeadroom ? " checked" : "") +
            "> Default to Showing Headroom</label>" +
            "</fieldset>" +
            "</div>";

        ensureVesselShapeParams(v);
        ensureVesselCustomTable(v);
        var tidCur = parseInt(v.vesselTypeId, 10) || 1;
        var isCustomType = tidCur === 15;
        var shapeH = v.shapeParams[0] || "25.00";
        var shapeW = v.shapeParams[1] || "10.50";
        var partOn = !!v.verticalSplitPartitioned;
        var partDis = !partOn;
        var splitShown = vesselTypeShowsVerticalSplitPartition(v.vesselTypeId || 1);
        var typePreviewMarkup = buildVesselShapePreviewMarkup(v.vesselTypeId || 1, shapeH, shapeW);

        var tabType =
            '<div class="vs-tab-panel" data-vs-panel="typeshape" role="tabpanel" hidden>' +
            '<div class="vs-ts-vessel-type-row">' +
            '<label class="vs-lbl" for="vs_vessel_type">Vessel Type:</label>' +
            '<select id="vs_vessel_type" class="vs-select vs-ts-vessel-dd">' +
            buildKeyedSelectOptions(VESSEL_TYPES, v.vesselTypeId, "id", "name") +
            "</select></div>" +
            '<div class="vs-ts-dist-row" id="vs_dist_row">Distance Units: <span id="vs_lbl_dist_val">' +
            escapeHtml(hLabel) +
            "</span></div>" +
            '<div class="vs-ts-main' +
            (isCustomType ? " vs-ts-main--custom" : "") +
            '">' +
            '<div class="vs-ts-left">' +
            '<div id="vs_shape_fields_slot">' +
            buildVesselShapeDimensionsHtml(v) +
            "</div>" +
            '<div class="vs-pnl-split"' +
            (splitShown ? "" : " hidden") +
            ' id="vs_pnl_split">' +
            '<label class="vs-check vs-split-chk">' +
            '<input type="checkbox" id="vs_partition_chk"' +
            (partOn ? " checked" : "") +
            '> Vertical Split / Partitioned</label>' +
            '<div class="vs-part-scale-row">' +
            '<label class="vs-lbl-part' +
            (partDis ? " vs-disabled" : "") +
            '" for="vs_partition_scale">Partition Scale:</label>' +
            '<input type="text" id="vs_partition_scale" class="vs-input vs-input-part"' +
            (partDis ? " disabled" : "") +
            ' value="' +
            escapeHtml(String(v.partitionScale || "")) +
            '">' +
            '<span class="vs-part-pct' +
            (partDis ? " vs-disabled" : "") +
            '">%</span>' +
            "</div></div></div>" +
            '<div class="vs-ts-diagram" id="vs_type_preview"' +
            (isCustomType ? " hidden" : "") +
            ">" +
            typePreviewMarkup +
            "</div></div>" +
            '<fieldset class="vs-group vs-assist-group">' +
            '<legend class="vs-group-legend">Assistance</legend>' +
            '<p id="vs_assistance" class="vs-assist-text vs-assist-centered">' +
            escapeHtml(VS_ASSIST_DEFAULT) +
            "</p>" +
            "</fieldset>" +
            "</div>";

        var densityChecked = v.densityMode !== "sg";
        var sgChecked = v.densityMode === "sg";
        var tabContents =
            '<div class="vs-tab-panel" data-vs-panel="contents" role="tabpanel" hidden>' +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Vessel Contents</legend>' +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_contents">Contents / Product:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_contents" class="vs-select vs-select-wide">' +
            buildVmSetupProductOptions(contents) +
            "</select></div></div>" +
            "</fieldset>" +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Density / Specific Gravity of Product</legend>' +
            '<label class="vs-radio-row">' +
            '<input type="radio" name="vs_den_mode" id="vs_den_density" value="density"' +
            (densityChecked ? " checked" : "") +
            '> Use Product Density:</label>' +
            '<div class="vs-indent">' +
            '<input type="text" id="vs_density_val" class="vs-input vs-input-sm" value="' +
            escapeHtml(String(v.productDensity || "")) +
            '">' +
            "</div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_density_units">Density Units:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_density_units" class="vs-select vs-select-wide">' +
            buildSelectOptionsStringList(DENSITY_UNITS_OPTIONS, v.densityUnits) +
            "</select></div></div>" +
            '<label class="vs-radio-row">' +
            '<input type="radio" name="vs_den_mode" id="vs_den_sg" value="sg"' +
            (sgChecked ? " checked" : "") +
            '> Use Specific Gravity:</label>' +
            '<div class="vs-indent">' +
            '<input type="text" id="vs_sg_val" class="vs-input vs-input-sm" value="' +
            escapeHtml(String(v.specificGravity || "")) +
            '">' +
            "</div>" +
            "</fieldset>" +
            "</div>";

        var stOpts = buildSensorTypeOptionsForNetwork(v.sensorNetworkId, v.sensorTypeId);
        var stId = stOpts.selectedId;
        /** SmartBob-II / SmartBob-II Average use pnlSmartBobA — Protocol B SmartBob uses pnlSensor (frmVesselSetup). */
        var showSmartBobPanel = stId === 1 || stId === 3;
        var showVegaPanel = stId === 11 || stId === 12;
        var showGenericPanel = !showSmartBobPanel && !showVegaPanel;
        var showGenericDecRow = stId === 4 || stId === 5;
        var showGenericDvRow = stId === 16;
        var showGenericOffsetRow = stId !== 10;
        var seOn = v.sensorEnabled !== false;
        var sbRowEn = v.sensorEnabled !== false;
        var sbEnMd = !!v.sbEnableMaxDrop;
        var sbAddr = escapeHtml(String(v.sensorAddress || ""));
        var sbOff = escapeHtml(String(v.sensorOffset || ""));
        var sbMax = escapeHtml(String(v.sbMaxDrop || ""));
        var sbMaxDisabled = !sbRowEn || !sbEnMd;
        /** frmVesselSetup PopulateSensorDecimalPlaces — items "1","2" only; default "2". */
        var decSel = String(v.sensorDecimalPlaces || "2");
        if (decSel !== "1" && decSel !== "2") decSel = "2";
        var decOpts = [1, 2]
            .map(function (n) {
                var ds = String(n);
                return (
                    '<option value="' +
                    ds +
                    '"' +
                    (decSel === ds ? " selected" : "") +
                    ">" +
                    ds +
                    "</option>"
                );
            })
            .join("");
        /** frmVesselSetup PopulateDistanceVariable — PV, SV, TV, QV; default SV. */
        var dvSel = String(v.sensorDistanceVariable || "SV").toUpperCase();
        if (["PV", "SV", "TV", "QV"].indexOf(dvSel) < 0) dvSel = "SV";
        var dvOpts = ["PV", "SV", "TV", "QV"]
            .map(function (L) {
                return (
                    '<option value="' +
                    L +
                    '"' +
                    (dvSel === L ? " selected" : "") +
                    ">" +
                    L +
                    "</option>"
                );
            })
            .join("");
        ensureVegaSensorsArray(v);
        var nVega = Math.min(32, Math.max(1, parseInt(v.vegaSensorCount || "1", 10) || 1));
        var vegaOffHdr = vegaOffsetHeaderText();
        var vegaTbodyInner = buildVegaSensorTbodyRowsHtml(v, nVega);
        var tabSensor =
            '<div class="vs-tab-panel" data-vs-panel="sensor" role="tabpanel" hidden>' +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_network">Sensor Network:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_network" class="vs-select vs-select-full">' +
            buildVmSetupNetworkOptions(v.sensorNetworkId) +
            "</select></div></div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_sensor_type">Sensor Type:</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_sensor_type" class="vs-select vs-select-wide">' +
            stOpts.html +
            "</select></div></div>" +
            '<div class="vs-sensor-stack">' +
            '<div id="vs_sensor_pnl_generic" class="vs-sensor-pnl"' +
            (showGenericPanel ? "" : " hidden") +
            ">" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_sensor_enabled">Enabled</label>' +
            '<div class="vs-field-control">' +
            '<input type="checkbox" id="vs_sensor_enabled"' +
            (seOn ? " checked" : "") +
            "></div></div>" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_sensor_address">Sensor Address</label>' +
            '<div class="vs-field-control">' +
            '<input type="text" id="vs_sensor_address" class="vs-input vs-input-addr" value="' +
            escapeHtml(String(v.sensorAddress || "")) +
            '"' +
            (seOn ? "" : " disabled") +
            "></div></div>" +
            '<div class="vs-field-row" id="vs_row_sensor_offset"' +
            (showGenericOffsetRow ? "" : " hidden") +
            ">" +
            '<label class="vs-lbl" for="vs_sensor_offset">Sensor Offset</label>' +
            '<div class="vs-field-control">' +
            '<input type="text" id="vs_sensor_offset" class="vs-input vs-input-offset" value="' +
            escapeHtml(String(v.sensorOffset || "")) +
            '"' +
            (seOn ? "" : " disabled") +
            '><span id="vs_sensor_offset_units" class="vs-static-units">' +
            escapeHtml(heightUnitsLabel()) +
            "</span></div></div>" +
            '<div class="vs-field-row" id="vs_row_decimal"' +
            (showGenericDecRow ? "" : " hidden") +
            ">" +
            '<label class="vs-lbl" for="vs_sensor_decimal">Decimal Places</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_sensor_decimal" class="vs-select vs-select-tiny">' +
            decOpts +
            "</select></div></div>" +
            '<div class="vs-field-row" id="vs_row_distvar"' +
            (showGenericDvRow ? "" : " hidden") +
            ">" +
            '<label class="vs-lbl" for="vs_sensor_distvar">Distance Variable</label>' +
            '<div class="vs-field-control">' +
            '<select id="vs_sensor_distvar" class="vs-select vs-select-tiny">' +
            dvOpts +
            "</select></div></div>" +
            "</div>" +
            '<div id="vs_sensor_pnl_smartbob" class="vs-sensor-pnl vs-sensor-pnl-smartbob"' +
            (showSmartBobPanel ? "" : " hidden") +
            ">" +
            '<table class="vs-sb-table vs-sb-single" aria-label="SmartBob sensor">' +
            "<thead><tr>" +
            '<th class="vs-sb-th">Enable</th>' +
            '<th class="vs-sb-th">Address</th>' +
            '<th class="vs-sb-th">Enable Max Drop</th>' +
            '<th class="vs-sb-th">Max Drop</th>' +
            '<th class="vs-sb-th">Sensor Offset</th>' +
            "</tr></thead><tbody><tr>" +
            '<td class="vs-sb-td vs-sb-td-chk"><input type="checkbox" id="vs_sb_en1"' +
            (sbRowEn ? " checked" : "") +
            "></td>" +
            '<td class="vs-sb-td"><input type="text" class="vs-input vs-sb-addr" id="vs_sb_addr1" value="' +
            sbAddr +
            '"></td>' +
            '<td class="vs-sb-td vs-sb-td-chk"><input type="checkbox" id="vs_sb_en_maxdrop"' +
            (sbEnMd ? " checked" : "") +
            (sbRowEn ? "" : " disabled") +
            "></td>" +
            '<td class="vs-sb-td"><input type="text" class="vs-input vs-sb-maxdrop" id="vs_sb_maxdrop" value="' +
            sbMax +
            '"' +
            (sbMaxDisabled ? " disabled" : "") +
            "></td>" +
            '<td class="vs-sb-td"><input type="text" class="vs-input vs-sb-offset" id="vs_sb_offset1" value="' +
            sbOff +
            '"' +
            (sbRowEn ? "" : " disabled") +
            "></td>" +
            "</tr></tbody></table></div>" +
            '<div id="vs_sensor_pnl_vega" class="vs-sensor-pnl vs-sensor-pnl-vega"' +
            (showVegaPanel ? "" : " hidden") +
            ">" +
            '<div class="vs-field-row">' +
            '<label class="vs-lbl" for="vs_vega_count">Sensor Count:</label>' +
            '<div class="vs-field-control">' +
            '<input type="number" id="vs_vega_count" class="vs-input vs-vega-count" min="1" max="32" step="1" value="' +
            escapeHtml(String(nVega)) +
            '">' +
            "</div></div>" +
            '<table class="vs-vega-table" aria-label="Vega sensors">' +
            "<thead><tr>" +
            '<th class="vs-vega-th-en">Enable</th>' +
            "<th>Address</th>" +
            "<th>" +
            escapeHtml(vegaOffHdr) +
            "</th>" +
            "<th>Distance Variable</th>" +
            "</tr></thead>" +
            '<tbody id="vs_vega_tbody">' +
            vegaTbodyInner +
            "</tbody></table>" +
            "</div></div></div>";

        var tabAlarms =
            '<div class="vs-tab-panel" data-vs-panel="alarms" role="tabpanel" hidden>' +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">High Level Alarms</legend>' +
            '<div class="vs-alarm-row">' +
            '<label class="vs-check vs-alarm-chk"><input type="checkbox" id="vs_ah_high"' +
            (v.alarmHighEnabled ? " checked" : "") +
            '> High Alarm Enabled</label>' +
            '<div class="vs-alarm-pct-cluster">' +
            '<input type="text" id="vs_ah_high_pct" class="vs-input vs-input-xs" value="' +
            escapeHtml(String(v.alarmHighPct || "")) +
            '">' +
            '<span class="vs-pct-lbl">% Full</span>' +
            "</div></div>" +
            '<div class="vs-alarm-row">' +
            '<label class="vs-check vs-alarm-chk"><input type="checkbox" id="vs_ah_prehigh"' +
            (v.alarmPreHighEnabled ? " checked" : "") +
            '> Pre-High Alarm Enabled</label>' +
            '<div class="vs-alarm-pct-cluster">' +
            '<input type="text" id="vs_ah_prehigh_pct" class="vs-input vs-input-xs" value="' +
            escapeHtml(String(v.alarmPreHighPct || "")) +
            '">' +
            '<span class="vs-pct-lbl">% Full</span>' +
            "</div></div>" +
            "</fieldset>" +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Low Level Alarms</legend>' +
            '<div class="vs-alarm-row">' +
            '<label class="vs-check vs-alarm-chk"><input type="checkbox" id="vs_ah_prelow"' +
            (v.alarmPreLowEnabled ? " checked" : "") +
            '> Pre-Low Alarm Enabled</label>' +
            '<div class="vs-alarm-pct-cluster">' +
            '<input type="text" id="vs_ah_prelow_pct" class="vs-input vs-input-xs" value="' +
            escapeHtml(String(v.alarmPreLowPct || "")) +
            '">' +
            '<span class="vs-pct-lbl">% Full</span>' +
            "</div></div>" +
            '<div class="vs-alarm-row">' +
            '<label class="vs-check vs-alarm-chk"><input type="checkbox" id="vs_ah_low"' +
            (v.alarmLowEnabled ? " checked" : "") +
            '> Low Alarm Enabled</label>' +
            '<div class="vs-alarm-pct-cluster">' +
            '<input type="text" id="vs_ah_low_pct" class="vs-input vs-input-xs" value="' +
            escapeHtml(String(v.alarmLowPct || "")) +
            '">' +
            '<span class="vs-pct-lbl">% Full</span>' +
            "</div></div>" +
            "</fieldset>" +
            "</div>";

        var ef = v.emailFlags || {};
        var simHidden =
            '<div class="sr-only" aria-hidden="true">' +
            '<input type="number" id="vs_h" step="0.01" value="' +
            v.heightFt +
            '">' +
            '<input type="number" id="vs_vol" step="1" value="' +
            v.volumeCuFt +
            '">' +
            '<input type="number" id="vs_w" step="1" value="' +
            v.weightLb +
            '">' +
            '<input type="number" id="vs_pct" min="0" max="100" step="1" value="' +
            Math.round(v.pctFull) +
            '">' +
            "</div>";

        var tabEmail =
            '<div class="vs-tab-panel" data-vs-panel="email" role="tabpanel" hidden>' +
            '<div class="vs-email-top">' +
            '<label class="vs-check vs-email-enable">' +
            '<input type="checkbox" id="vs_email_en"' +
            (v.emailNotificationsEnabled ? " checked" : "") +
            '> Enable Email Notifications</label>' +
            '<button type="button" class="vs-btn" id="vs_sel_contacts" disabled>Select Contacts</button>' +
            "</div>" +
            '<div class="vs-table-wrap vs-email-grid-wrap">' +
            '<table class="vs-data-grid">' +
            "<thead><tr><th>First Name</th><th>Last Name</th><th>Email</th></tr></thead>" +
            '<tbody id="vs_email_contacts_tbody">' +
            (state.contacts && state.contacts.length
                ? buildVesselAssignedContactRowsHtml(v)
                : '<tr><td colspan="3" class="vs-muted">No contacts — add them in Contact Maintenance.</td></tr>') +
            "</tbody></table></div>" +
            '<fieldset class="vs-group">' +
            '<legend class="vs-group-legend">Email Notifications</legend>' +
            '<div class="vs-email-grid">' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_high"' +
            (ef.high ? " checked" : "") +
            '> High Alarm Enabled</label>' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_prehigh"' +
            (ef.preHigh ? " checked" : "") +
            '> Pre-High Alarm Enabled</label>' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_prelow"' +
            (ef.preLow ? " checked" : "") +
            '> Pre-Low Alarm Enabled</label>' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_low"' +
            (ef.low ? " checked" : "") +
            '> Low Alarm Enabled</label>' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_status"' +
            (ef.vesselStatus ? " checked" : "") +
            '> Vessel Status Enabled</label>' +
            '<label class="vs-check"><input type="checkbox" id="vs_ef_err"' +
            (ef.error ? " checked" : "") +
            '> Error Alarm Enabled</label>' +
            "</div></fieldset></div>";

        return (
            '<div class="vs-shell">' +
            '<div class="vs-inner-title" id="vs_lbl_title">Vessel Setup - ' +
            escapeHtml(nm) +
            "</div>" +
            tabStrip +
            '<div class="vs-tab-panels">' +
            tabDisplay +
            tabType +
            tabContents +
            tabSensor +
            tabAlarms +
            tabEmail +
            "</div>" +
            simHidden +
            "</div>"
        );
    }

    function bindVesselSetupForm() {
        var shell = appModalBody.querySelector(".vs-shell");
        if (!shell) return;

        var tabs = shell.querySelectorAll(".vs-tab");
        var panels = shell.querySelectorAll(".vs-tab-panel");

        function showTab(key) {
            tabs.forEach(function (t) {
                var on = t.getAttribute("data-vs-tab") === key;
                t.classList.toggle("active", on);
                t.setAttribute("aria-selected", on ? "true" : "false");
            });
            panels.forEach(function (p) {
                var on = p.getAttribute("data-vs-panel") === key;
                p.classList.toggle("active", on);
                if (on) p.removeAttribute("hidden");
                else p.setAttribute("hidden", "");
            });
        }

        tabs.forEach(function (btn) {
            btn.addEventListener("click", function () {
                showTab(btn.getAttribute("data-vs-tab"));
            });
        });

        var nameEl = document.getElementById("vs_name");
        var titleEl = document.getElementById("vs_lbl_title");
        if (nameEl && titleEl) {
            nameEl.addEventListener("input", function () {
                titleEl.textContent = "Vessel Setup - " + nameEl.value;
            });
        }

        var colorEl = document.getElementById("vs_color");
        var prevEl = document.getElementById("vs_color_preview");
        if (colorEl && prevEl) {
            colorEl.addEventListener("change", function () {
                prevEl.style.background = fillColorPreviewCss(colorEl.value);
            });
        }

        var typeEl = document.getElementById("vs_vessel_type");
        var asstEl = document.getElementById("vs_assistance");
        var previewEl = document.getElementById("vs_type_preview");
        var shapeFieldsSlot = document.getElementById("vs_shape_fields_slot");
        var partChk = document.getElementById("vs_partition_chk");
        var partScale = document.getElementById("vs_partition_scale");
        var partPct = document.querySelector(".vs-part-pct");
        var partLbl = document.querySelector(".vs-lbl-part");

        function syncTypeShapeDiagram() {
            if (!previewEl || !typeEl) return;
            var tid = parseInt(typeEl.value, 10) || 1;
            if (tid === 15) {
                previewEl.innerHTML = "";
                return;
            }
            var hEl = document.getElementById("vs_sp_0");
            var wEl = document.getElementById("vs_sp_1");
            var dh = hEl ? hEl.value : "";
            var dw = wEl ? wEl.value : "";
            previewEl.innerHTML = buildVesselShapePreviewMarkup(tid, dh, dw);
        }

        function syncCustomLayoutVisibility() {
            var main = shell.querySelector(".vs-ts-main");
            if (!main || !typeEl) return;
            var tid = parseInt(typeEl.value, 10) || 1;
            if (tid === 15) {
                main.classList.add("vs-ts-main--custom");
                if (previewEl) previewEl.setAttribute("hidden", "");
            } else {
                main.classList.remove("vs-ts-main--custom");
                if (previewEl) previewEl.removeAttribute("hidden");
            }
        }

        function syncPartitionPanelVisibility() {
            var pnl = document.getElementById("vs_pnl_split");
            if (!pnl || !typeEl) return;
            var tid = parseInt(typeEl.value, 10) || 1;
            if (vesselTypeShowsVerticalSplitPartition(tid)) {
                pnl.removeAttribute("hidden");
            } else {
                pnl.setAttribute("hidden", "");
            }
        }

        function syncPartitionControls() {
            var on = partChk && partChk.checked;
            if (partScale) {
                partScale.disabled = !on;
            }
            if (partLbl) partLbl.classList.toggle("vs-disabled", !on);
            if (partPct) partPct.classList.toggle("vs-disabled", !on);
        }

        function setAssistance(text) {
            if (asstEl) asstEl.textContent = text;
        }

        if (typeEl && asstEl) {
            typeEl.addEventListener("change", function () {
                var vv = findVessel(vmSetupEditingId);
                var newTid = parseInt(typeEl.value, 10) || 1;
                if (vv) {
                    var prevTid = parseInt(vv.vesselTypeId, 10) || 1;
                    if (prevTid === 15) {
                        refreshVesselCustomStrapFromDom(vv);
                    } else {
                        refreshVesselShapeParamsFromDom(vv, prevTid);
                    }
                    vv.vesselTypeId = newTid;
                }
                setAssistance(assistanceTextForVesselType(newTid) || VS_ASSIST_DEFAULT);
                if (shapeFieldsSlot && vv) {
                    ensureVesselShapeParams(vv);
                    if (newTid === 15) {
                        ensureVesselCustomTable(vv);
                    }
                    shapeFieldsSlot.innerHTML = buildVesselShapeDimensionsHtml(vv);
                }
                syncCustomLayoutVisibility();
                syncTypeShapeDiagram();
                syncPartitionPanelVisibility();
            });
            typeEl.addEventListener("mouseenter", function () {
                setAssistance("Use this control to select the type or shape of the vessel.");
            });
        }

        if (shapeFieldsSlot) {
            shapeFieldsSlot.addEventListener("input", syncTypeShapeDiagram);
            shapeFieldsSlot.addEventListener(
                "blur",
                function (ev) {
                    if (ev.target && ev.target.classList && ev.target.classList.contains("vs-custom-d")) {
                        sortCustomStrapTbodyByDistance();
                    }
                },
                true
            );
            shapeFieldsSlot.addEventListener(
                "mouseenter",
                function (ev) {
                    var t = ev.target;
                    if (!t || !t.classList) return;
                    if (t.classList.contains("vs-shape-param")) {
                        var u = document.getElementById("vs_lbl_dist_val");
                        var ut = u ? u.textContent : "";
                        setAssistance(
                            "Enter this dimension in " +
                            ut.toLowerCase() +
                            ". Distance Units are set under Site Maintenance."
                        );
                    }
                    if (t.classList.contains("vs-custom-d") || t.classList.contains("vs-custom-o")) {
                        setAssistance(
                            "A valid custom table must have at least two rows or points in it. Rows are sorted by distance."
                        );
                    }
                },
                true
            );
        }

        if (!shell._ebobVsStrapBound) {
            shell._ebobVsStrapBound = true;
            shell.addEventListener("click", function (e) {
                var id = e.target && e.target.id;
                var tbSel = document.getElementById("vs_custom_tbody");
                if (
                    tbSel &&
                    tbSel.contains(e.target) &&
                    !e.target.closest("button") &&
                    e.target.closest("tr.vs-custom-data-row")
                ) {
                    var trHit = e.target.closest("tr.vs-custom-data-row");
                    if (trHit) {
                        tbSel.querySelectorAll("tr.vs-custom-data-row").forEach(function (r) {
                            r.classList.remove("vs-custom-row-selected");
                        });
                        trHit.classList.add("vs-custom-row-selected");
                    }
                }
                if (id === "vs_custom_add") {
                    e.preventDefault();
                    var tb = document.getElementById("vs_custom_tbody");
                    if (tb) {
                        var tr = document.createElement("tr");
                        tr.className = "vs-custom-data-row";
                        tr.innerHTML =
                            '<td><input type="text" class="vs-input vs-input-custom vs-custom-d" value=""></td>' +
                            '<td><input type="text" class="vs-input vs-input-custom vs-custom-o" value=""></td>';
                        tb.appendChild(tr);
                        tb.querySelectorAll("tr.vs-custom-data-row").forEach(function (r) {
                            r.classList.remove("vs-custom-row-selected");
                        });
                        tr.classList.add("vs-custom-row-selected");
                        toast("Row added.");
                    }
                } else if (id === "vs_custom_delete") {
                    e.preventDefault();
                    var tb2 = document.getElementById("vs_custom_tbody");
                    if (!tb2) return;
                    var sel = tb2.querySelector("tr.vs-custom-data-row.vs-custom-row-selected");
                    if (sel) {
                        sel.remove();
                        toast("Row removed.");
                    } else {
                        toast("Select a row to delete (click the row).");
                    }
                } else if (id === "vs_custom_export") {
                    e.preventDefault();
                    var data = readCustomStrapRowsFromDom();
                    if (!data || !data.length) {
                        showBinventoryMessageBox({
                            icon: "warn",
                            message: "There is no table data to export.",
                            buttons: "ok"
                        });
                        return;
                    }
                    showBinventoryMessageBox({
                        icon: "info",
                        message:
                            "The custom/lookup table was successfully copied to the Windows clipboard.",
                        buttons: "ok"
                    });
                } else if (id === "vs_custom_import") {
                    e.preventDefault();
                    var tbImp = document.getElementById("vs_custom_tbody");
                    var rowCount = tbImp ? tbImp.querySelectorAll("tr.vs-custom-data-row").length : 0;
                    function openVsImportPasteBackdrop() {
                        var ta = document.getElementById("vsImportPasteTa");
                        var bd = document.getElementById("backdropVsImportPaste");
                        if (ta) ta.value = "";
                        if (bd) bd.classList.add("show");
                    }
                    if (rowCount > 0) {
                        showBinventoryMessageBox({
                            icon: "question",
                            message:
                                "Importing will overwrite the existing table.  Are you sure you want to continue?",
                            buttons: "okcancel",
                            onOk: openVsImportPasteBackdrop
                        });
                    } else {
                        openVsImportPasteBackdrop();
                    }
                }
            });
            shell.addEventListener("change", function (e) {
                var t = e.target;
                if (!t || t.id !== "vs_custom_output_type") return;
                var sel = t;
                var idx = parseInt(sel.value, 10) || 0;
                var da = sel.getAttribute("data-dist-abbr") || "ft";
                var va = sel.getAttribute("data-vol-abbr") || "gal";
                var wa = sel.getAttribute("data-wt-abbr") || "tons";
                var th = document.getElementById("vs_custom_th_out");
                if (th) {
                    th.textContent = customOutputColumnHeader(idx, da, va, wa);
                }
            });
        }

        var distRow = document.getElementById("vs_dist_row");
        if (distRow) {
            distRow.addEventListener("mouseenter", function () {
                var u = document.getElementById("vs_lbl_dist_val");
                var ut = u ? u.textContent : "";
                setAssistance(
                    "Enter all vessel dimensions in " +
                    ut.toLowerCase() +
                    ". Distance Units are set under Site Maintenance."
                );
            });
        }

        if (partChk) {
            partChk.addEventListener("change", function () {
                syncPartitionControls();
                setAssistance(
                    partChk.checked
                        ? "Check this to make a vessel partition or vertical split."
                        : VS_ASSIST_DEFAULT
                );
            });
            partChk.addEventListener("mouseenter", function () {
                setAssistance("Check this to make a vessel partition or vertical split.");
            });
            syncPartitionControls();
        }
        if (partScale) {
            partScale.addEventListener("mouseenter", function () {
                setAssistance(
                    "Enter the size or scale of the partition as a percentage of the whole vessel. Common partition scales are 25, 33, 50 and 66 percent."
                );
            });
        }

        var tsPanel = shell.querySelector('[data-vs-panel="typeshape"]');
        if (tsPanel && asstEl) {
            tsPanel.addEventListener("mouseleave", function () {
                setAssistance(VS_ASSIST_DEFAULT);
            });
        }

        syncCustomLayoutVisibility();
        syncTypeShapeDiagram();
        syncPartitionPanelVisibility();

        var emailEn = document.getElementById("vs_email_en");
        var grpEmail = shell.querySelector(".vs-email-grid");
        var selBtn = document.getElementById("vs_sel_contacts");
        function syncEmail() {
            var on = emailEn && emailEn.checked;
            if (grpEmail) grpEmail.style.opacity = on ? "1" : "0.5";
            if (selBtn) selBtn.disabled = !on;
            shell.querySelectorAll(".vs-email-grid input").forEach(function (inp) {
                inp.disabled = !on;
            });
        }
        if (emailEn) {
            emailEn.addEventListener("change", syncEmail);
            syncEmail();
        }
        if (selBtn) {
            selBtn.addEventListener("click", function () {
                if (!emailEn || !emailEn.checked) return;
                openAssignContactsDialog();
            });
        }

        var sensorTypeEl = document.getElementById("vs_sensor_type");
        var networkEl = document.getElementById("vs_network");

        function syncSbMaxDrop() {
            var rowOn = document.getElementById("vs_sb_en1") && document.getElementById("vs_sb_en1").checked;
            var mdEn = document.getElementById("vs_sb_en_maxdrop") && document.getElementById("vs_sb_en_maxdrop").checked;
            var md = document.getElementById("vs_sb_maxdrop");
            var em = document.getElementById("vs_sb_en_maxdrop");
            if (em) em.disabled = !rowOn;
            if (md) md.disabled = !rowOn || !mdEn;
        }

        function syncSbRowEnabled() {
            var en = document.getElementById("vs_sb_en1");
            var addr = document.getElementById("vs_sb_addr1");
            var off = document.getElementById("vs_sb_offset1");
            var on = en && en.checked;
            if (addr) addr.disabled = !on;
            if (off) off.disabled = !on;
            syncSbMaxDrop();
        }

        function applySensorTypeOptionsFromNetwork() {
            if (!networkEl || !sensorTypeEl) return;
            var nid = networkEl.value;
            var cur = parseInt(sensorTypeEl.value, 10);
            var out = buildSensorTypeOptionsForNetwork(nid, cur);
            sensorTypeEl.innerHTML = out.html;
            sensorTypeEl.value = String(out.selectedId);
            syncSensorPanels();
        }

        function syncSensorPanels() {
            if (!sensorTypeEl) return;
            var st = parseInt(sensorTypeEl.value, 10) || 1;
            var gen = document.getElementById("vs_sensor_pnl_generic");
            var sb = document.getElementById("vs_sensor_pnl_smartbob");
            var vg = document.getElementById("vs_sensor_pnl_vega");
            var showSmart = st === 1 || st === 3;
            var showVega = st === 11 || st === 12;
            if (gen) {
                if (showSmart || showVega) gen.setAttribute("hidden", "");
                else gen.removeAttribute("hidden");
            }
            if (sb) {
                if (showSmart) sb.removeAttribute("hidden");
                else sb.setAttribute("hidden", "");
            }
            if (vg) {
                if (showVega) vg.removeAttribute("hidden");
                else vg.setAttribute("hidden", "");
            }
            var rowDec = document.getElementById("vs_row_decimal");
            var rowDv = document.getElementById("vs_row_distvar");
            var rowOff = document.getElementById("vs_row_sensor_offset");
            if (rowDec && rowDv && rowOff && !showSmart && !showVega) {
                var showDec = st === 4 || st === 5;
                var showDv = st === 16;
                var showOff = st !== 10;
                if (showDv) {
                    rowDec.setAttribute("hidden", "");
                    rowDv.removeAttribute("hidden");
                } else {
                    rowDv.setAttribute("hidden", "");
                    if (showDec) rowDec.removeAttribute("hidden");
                    else rowDec.setAttribute("hidden", "");
                }
                if (showOff) rowOff.removeAttribute("hidden");
                else rowOff.setAttribute("hidden", "");
            }
            if (showSmart) {
                syncSbRowEnabled();
            }
        }

        function syncSensorAddrEnabled() {
            var en = document.getElementById("vs_sensor_enabled");
            var addr = document.getElementById("vs_sensor_address");
            var off = document.getElementById("vs_sensor_offset");
            var on = en && en.checked;
            if (addr) addr.disabled = !on;
            if (off) off.disabled = !on;
        }

        var sensorEnChk = document.getElementById("vs_sensor_enabled");
        if (sensorEnChk) {
            sensorEnChk.addEventListener("change", syncSensorAddrEnabled);
            syncSensorAddrEnabled();
        }
        if (networkEl) {
            networkEl.addEventListener("change", applySensorTypeOptionsFromNetwork);
        }
        if (sensorTypeEl) {
            sensorTypeEl.addEventListener("change", syncSensorPanels);
            syncSensorPanels();
        }

        var sbEn1 = document.getElementById("vs_sb_en1");
        var sbEnMax = document.getElementById("vs_sb_en_maxdrop");
        if (sbEn1) {
            sbEn1.addEventListener("change", syncSbRowEnabled);
        }
        if (sbEnMax) {
            sbEnMax.addEventListener("change", syncSbMaxDrop);
        }
        if (sbEn1 || sbEnMax) {
            syncSbRowEnabled();
        }

        function snapshotVegaRowsIntoVessel() {
            var vv = findVessel(vmSetupEditingId);
            if (!vv) return;
            ensureVegaSensorsArray(vv);
            var ix;
            for (ix = 0; ix < 32; ix++) {
                var enG = document.getElementById("vs_vg_en" + (ix + 1));
                var aG = document.getElementById("vs_vg_a" + (ix + 1));
                var oG = document.getElementById("vs_vg_o" + (ix + 1));
                var dG = document.getElementById("vs_vg_dv" + (ix + 1));
                if (enG) vv.vegaSensors[ix].enabled = enG.checked;
                if (aG) vv.vegaSensors[ix].address = aG.value;
                if (oG) vv.vegaSensors[ix].offset = oG.value;
                if (dG) vv.vegaSensors[ix].dv = dG.value;
            }
        }

        function rebuildVegaSensorTable() {
            var vv = findVessel(vmSetupEditingId);
            if (!vv) return;
            var vcEl = document.getElementById("vs_vega_count");
            var tbody = document.getElementById("vs_vega_tbody");
            if (!vcEl || !tbody) return;
            var prevTr = tbody.querySelectorAll("tr").length;
            var prevN =
                prevTr > 0
                    ? prevTr
                    : Math.min(32, Math.max(1, parseInt(vv.vegaSensorCount || "1", 10) || 1));
            var n = Math.min(32, Math.max(1, parseInt(vcEl.value, 10) || 1));
            snapshotVegaRowsIntoVessel();
            if (n > prevN) {
                ensureVegaSensorsArray(vv);
                var j;
                for (j = prevN; j < n; j++) {
                    vv.vegaSensors[j].enabled = true;
                    if (!vv.vegaSensors[j].address || String(vv.vegaSensors[j].address).trim() === "") {
                        vv.vegaSensors[j].address = String(j + 1);
                    }
                    if (vv.vegaSensors[j].offset == null || String(vv.vegaSensors[j].offset).trim() === "") {
                        vv.vegaSensors[j].offset = "0.00";
                    }
                    vv.vegaSensors[j].dv = vv.vegaSensors[j].dv || "SV";
                }
            }
            vv.vegaSensorCount = String(n);
            vcEl.value = String(n);
            tbody.innerHTML = buildVegaSensorTbodyRowsHtml(vv, n);
            bindVegaRowEnableHandlers();
        }

        function bindVegaRowEnableHandlers() {
            var vcEl = document.getElementById("vs_vega_count");
            var nR = vcEl ? Math.min(32, Math.max(1, parseInt(vcEl.value, 10) || 1)) : 1;
            var r;
            for (r = 1; r <= nR; r++) {
                (function (rowIdx) {
                    var enR = document.getElementById("vs_vg_en" + rowIdx);
                    var aR = document.getElementById("vs_vg_a" + rowIdx);
                    var oR = document.getElementById("vs_vg_o" + rowIdx);
                    var dR = document.getElementById("vs_vg_dv" + rowIdx);
                    function syncVegaRow() {
                        var on = enR && enR.checked;
                        if (aR) aR.disabled = !on;
                        if (oR) oR.disabled = !on;
                        if (dR) dR.disabled = !on;
                    }
                    if (enR) {
                        enR.addEventListener("change", syncVegaRow);
                        syncVegaRow();
                    }
                })(r);
            }
        }

        var vegaCountEl = document.getElementById("vs_vega_count");
        if (vegaCountEl) {
            vegaCountEl.addEventListener("change", rebuildVegaSensorTable);
        }
        bindVegaRowEnableHandlers();
    }

    function closeVesselSetupToMaintenance() {
        dismissAssignContactsBackdrop();
        appModalShell.classList.remove("modal-vessel-setup");
        appModalShell.classList.remove("modal-sn-networks");
        appModalShell.classList.remove("modal-system-setup");
        appModalShell.classList.remove("modal-site-maintenance");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-report-preview");
        appModalShell.classList.remove("modal-report-crystal");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-user-maintenance");
        vmSetupEditingId = null;
        updateAppModalContent(
            "Vessel Maintenance — Binventory Workstation",
            buildVmListHtml(),
            ""
        );
        bindVmList();
        vmRefreshToolbar();
    }

    function openVmSetupView(vesselId) {
        var v = findVessel(vesselId);
        if (!v) return;
        ensureVesselFields(v, 0);
        vesselSetupSnapshotById[vesselId] = takeVesselSetupSnapshot(v);
        vmSetupEditingId = vesselId;
        appModalShell.classList.remove("modal-sn-networks");
        appModalShell.classList.remove("modal-site-maintenance");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-system-setup");
        appModalShell.classList.remove("modal-report-preview");
        appModalShell.classList.remove("modal-report-crystal");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-user-maintenance");
        appModalShell.classList.add("modal-vessel-setup");
        updateAppModalContent(
            "Vessel Setup — Binventory Workstation",
            buildVesselSetupBody(v),
            '<button type="button" class="primary" id="vsSave">Save</button><button type="button" class="secondary" id="vsCancel">Cancel</button>'
        );
        bindVesselSetupForm();

        document.getElementById("vsSave").addEventListener("click", function () {
            var vv = findVessel(vmSetupEditingId);
            if (!vv) return;
            var sensorErr = validateVesselSetupSensorBlockFromDom();
            if (!sensorErr.ok) {
                toast(sensorErr.message);
                return;
            }
            if (assignContactsState && assignContactsState.vesselId === vmSetupEditingId) {
                ensureVesselContactIds(vv);
                vv.vesselContactIds = sortContactIdsStable(assignContactsState.assigned.slice());
                dismissAssignContactsBackdrop();
                refreshVesselEmailContactsTable();
            }
            vv.name = (document.getElementById("vs_name") && document.getElementById("vs_name").value.trim()) || vv.name;
            vv.fillColor = document.getElementById("vs_color") ? document.getElementById("vs_color").value : vv.fillColor;
            vv.volumeDisplayUnits = document.getElementById("vs_vol_units")
                ? document.getElementById("vs_vol_units").value
                : vv.volumeDisplayUnits;
            vv.weightDisplayUnits = document.getElementById("vs_wt_units")
                ? document.getElementById("vs_wt_units").value
                : vv.weightDisplayUnits;
            vv.defaultHeadroom = document.getElementById("vs_headroom")
                ? document.getElementById("vs_headroom").checked
                : false;
            vv.headroom = vv.defaultHeadroom;
            vv.vesselTypeId = parseInt(document.getElementById("vs_vessel_type").value, 10) || 1;
            refreshVesselShapeParamsFromDom(vv);
            if (document.getElementById("vs_partition_chk")) {
                vv.verticalSplitPartitioned = document.getElementById("vs_partition_chk").checked;
            }
            if (document.getElementById("vs_partition_scale")) {
                vv.partitionScale = document.getElementById("vs_partition_scale").value;
            }
            if (!vesselTypeShowsVerticalSplitPartition(vv.vesselTypeId)) {
                vv.verticalSplitPartitioned = false;
                vv.partitionScale = "100";
            }
            var cSel = document.getElementById("vs_contents");
            vv.contents = cSel ? cSel.value.trim() || vv.contents : vv.contents;
            vv.product = vv.contents;
            vv.densityMode = document.getElementById("vs_den_sg") && document.getElementById("vs_den_sg").checked ? "sg" : "density";
            vv.productDensity = document.getElementById("vs_density_val")
                ? document.getElementById("vs_density_val").value
                : vv.productDensity;
            vv.densityUnits = document.getElementById("vs_density_units")
                ? document.getElementById("vs_density_units").value
                : vv.densityUnits;
            vv.specificGravity = document.getElementById("vs_sg_val")
                ? document.getElementById("vs_sg_val").value
                : vv.specificGravity;
            vv.heightFt = parseFloat(document.getElementById("vs_h").value) || vv.heightFt;
            vv.volumeCuFt = parseInt(document.getElementById("vs_vol").value, 10) || vv.volumeCuFt;
            vv.weightLb = parseInt(document.getElementById("vs_w").value, 10) || vv.weightLb;
            vv.pctFull = Math.min(
                100,
                Math.max(0, parseInt(document.getElementById("vs_pct").value, 10) || vv.pctFull)
            );
            vv.sensorNetworkId = document.getElementById("vs_network")
                ? document.getElementById("vs_network").value || vv.sensorNetworkId
                : vv.sensorNetworkId;
            vv.sensorTypeId = parseInt(document.getElementById("vs_sensor_type").value, 10) || 1;
            var stSave = vv.sensorTypeId;
            var bSave = getSensorAddressBounds(stSave);
            if (stSave === 1 || stSave === 3) {
                if (document.getElementById("vs_sb_en1")) {
                    vv.sensorEnabled = document.getElementById("vs_sb_en1").checked;
                }
                if (document.getElementById("vs_sb_addr1")) {
                    vv.sensorAddress = validateIntegerSensorAddress(
                        document.getElementById("vs_sb_addr1").value,
                        bSave.min,
                        bSave.max
                    ).normalized;
                }
                if (document.getElementById("vs_sb_offset1")) {
                    vv.sensorOffset = document.getElementById("vs_sb_offset1").value;
                }
                if (document.getElementById("vs_sb_en_maxdrop")) {
                    vv.sbEnableMaxDrop = document.getElementById("vs_sb_en_maxdrop").checked;
                }
                if (document.getElementById("vs_sb_maxdrop")) {
                    vv.sbMaxDrop = document.getElementById("vs_sb_maxdrop").value;
                }
            } else if (stSave === 11 || stSave === 12) {
                ensureVegaSensorsArray(vv);
                var vcSg = document.getElementById("vs_vega_count");
                var nSg = Math.min(32, Math.max(1, parseInt(vcSg && vcSg.value, 10) || 1));
                vv.vegaSensorCount = String(nSg);
                var vx;
                for (vx = 0; vx < 32; vx++) {
                    var enSg = document.getElementById("vs_vg_en" + (vx + 1));
                    var aSg = document.getElementById("vs_vg_a" + (vx + 1));
                    var oSg = document.getElementById("vs_vg_o" + (vx + 1));
                    var dSg = document.getElementById("vs_vg_dv" + (vx + 1));
                    if (!vv.vegaSensors[vx]) vv.vegaSensors[vx] = {};
                    if (enSg) vv.vegaSensors[vx].enabled = enSg.checked;
                    if (aSg) {
                        vv.vegaSensors[vx].address = validateIntegerSensorAddress(
                            aSg.value,
                            bSave.min,
                            bSave.max
                        ).normalized;
                    }
                    if (oSg) vv.vegaSensors[vx].offset = oSg.value;
                    if (dSg) vv.vegaSensors[vx].dv = dSg.value;
                }
                if (vv.vegaSensors[0]) {
                    vv.sensorEnabled = vv.vegaSensors[0].enabled;
                    vv.sensorAddress = vv.vegaSensors[0].address || "";
                    vv.sensorOffset = vv.vegaSensors[0].offset || "";
                }
            } else {
                if (document.getElementById("vs_sensor_enabled")) {
                    vv.sensorEnabled = document.getElementById("vs_sensor_enabled").checked;
                }
                if (document.getElementById("vs_sensor_address")) {
                    vv.sensorAddress = validateIntegerSensorAddress(
                        document.getElementById("vs_sensor_address").value,
                        bSave.min,
                        bSave.max
                    ).normalized;
                }
                if (document.getElementById("vs_sensor_offset")) {
                    vv.sensorOffset = document.getElementById("vs_sensor_offset").value;
                }
            }
            if (document.getElementById("vs_sensor_decimal")) {
                vv.sensorDecimalPlaces = document.getElementById("vs_sensor_decimal").value;
            }
            if (document.getElementById("vs_sensor_distvar")) {
                vv.sensorDistanceVariable = document.getElementById("vs_sensor_distvar").value;
            }
            vv.alarmHighEnabled = document.getElementById("vs_ah_high").checked;
            vv.alarmHighPct = document.getElementById("vs_ah_high_pct").value;
            vv.alarmPreHighEnabled = document.getElementById("vs_ah_prehigh").checked;
            vv.alarmPreHighPct = document.getElementById("vs_ah_prehigh_pct").value;
            vv.alarmPreLowEnabled = document.getElementById("vs_ah_prelow").checked;
            vv.alarmPreLowPct = document.getElementById("vs_ah_prelow_pct").value;
            vv.alarmLowEnabled = document.getElementById("vs_ah_low").checked;
            vv.alarmLowPct = document.getElementById("vs_ah_low_pct").value;
            vv.emailNotificationsEnabled = document.getElementById("vs_email_en").checked;
            vv.emailFlags = {
                high: document.getElementById("vs_ef_high").checked,
                preHigh: document.getElementById("vs_ef_prehigh").checked,
                preLow: document.getElementById("vs_ef_prelow").checked,
                low: document.getElementById("vs_ef_low").checked,
                vesselStatus: document.getElementById("vs_ef_status").checked,
                error: document.getElementById("vs_ef_err").checked
            };
            saveState();
            delete vesselSetupSnapshotById[vmSetupEditingId];
            closeVesselSetupToMaintenance();
            refreshUI();
            toast("Vessel saved (simulation).");
        });
        document.getElementById("vsCancel").addEventListener("click", function () {
            closeVesselSetupDiscardChanges();
        });
    }

    function openVesselMaintenance() {
        state.vessels.forEach(function (v, idx) {
            ensureVesselFields(v, idx);
        });
        var firstSorted = vesselsSortedList()[0];
        vmSelectedId = firstSorted ? firstSorted.id : null;
        openAppModal(
            "Vessel Maintenance — Binventory Workstation",
            buildVmListHtml(),
            ""
        );
        bindVmList();
    }

    function parseGroupNumericId(g) {
        var n = parseInt(String((g && g.id) || "").replace(/^g/i, ""), 10);
        return isNaN(n) ? 0 : n;
    }

    function nextGroupRowId() {
        var max = 0;
        state.groups.forEach(function (g) {
            var n = parseGroupNumericId(g);
            if (n > max) max = n;
        });
        return "g" + (max + 1);
    }

    function findGroupById(id) {
        return state.groups.filter(function (g) {
            return g.id === id;
        })[0];
    }

    function groupsListSorted() {
        var list = groupsForWorkstationSite().slice();
        list.sort(function (a, b) {
            return parseGroupNumericId(a) - parseGroupNumericId(b);
        });
        return list;
    }

    function removeGroupIdFromSchedules(groupId) {
        state.schedules.forEach(function (s) {
            if (!s.groupIds || !s.groupIds.length) return;
            s.groupIds = s.groupIds.filter(function (gid) {
                return gid !== groupId;
            });
        });
    }

    /** Mirrors giAccessLevel — Administrator/SystemUser vs ReadOnly; logged out → ReadOnly (Logoff_Click). */
    function currentUserAccessLevel() {
        if (!isSessionLoggedIn()) return "readonly";
        var u = state.users.filter(function (x) {
            return x.name === state.currentUser;
        })[0];
        if (u && u.role === "Read Only") return "readonly";
        return "admin";
    }

    /**
     * frmVesselGroup — Assign Vessels to Group (584×439). Mutates working.vesselIds; onClose returns to Group Setup.
     */
    function openVesselGroupAssignDialog(working, onClose) {
        var selectedUnId = null;
        var selectedAsId = null;

        function unassignedVessels() {
            var assigned = {};
            working.vesselIds.forEach(function (vid) {
                assigned[vid] = true;
            });
            return vesselsForCurrentSiteSorted().filter(function (v) {
                return !assigned[v.id];
            });
        }

        function assignedVesselsOrdered() {
            return working.vesselIds
                .map(function (vid) {
                    return findVessel(vid);
                })
                .filter(Boolean);
        }

        function renderVgaTables() {
            var unBody = document.getElementById("vgaUnBody");
            var asBody = document.getElementById("vgaAsBody");
            if (!unBody || !asBody) return;
            unBody.innerHTML = unassignedVessels()
                .map(function (v) {
                    var sel = v.id === selectedUnId ? " selected" : "";
                    return (
                        '<tr class="vga-row vga-un-row' +
                        sel +
                        '" data-vessel-id="' +
                        escapeHtml(v.id) +
                        '">' +
                        '<td class="vga-col-id">' +
                        escapeHtml(v.id) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.name) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.contents || "") +
                        "</td></tr>"
                    );
                })
                .join("");
            asBody.innerHTML = assignedVesselsOrdered()
                .map(function (v) {
                    var sel = v.id === selectedAsId ? " selected" : "";
                    return (
                        '<tr class="vga-row vga-as-row' +
                        sel +
                        '" data-vessel-id="' +
                        escapeHtml(v.id) +
                        '">' +
                        '<td class="vga-col-id">' +
                        escapeHtml(v.id) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.name) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.contents || "") +
                        "</td></tr>"
                    );
                })
                .join("");
        }

        var html =
            '<div class="vga-shell">' +
            '<div class="vga-lbl-title">Assign Vessels to Group</div>' +
            '<fieldset class="win-groupbox vga-gb"><legend>Unassigned Vessels</legend>' +
            '<div class="vga-dgv-wrap">' +
            '<table class="win-dgv vga-dgv" id="vgaUnTable">' +
            "<thead><tr>" +
            '<th class="vga-col-id">VesselID</th>' +
            "<th>Vessel Name</th>" +
            "<th>Contents</th>" +
            "</tr></thead>" +
            '<tbody id="vgaUnBody"></tbody>' +
            "</table></div></fieldset>" +
            '<div class="vga-btn-row">' +
            '<button type="button" class="win-btn" id="vgaAdd">Add</button>' +
            '<button type="button" class="win-btn" id="vgaRemove">Remove</button>' +
            '<button type="button" class="win-btn" id="vgaAddAll">Add All</button>' +
            '<button type="button" class="win-btn" id="vgaRemoveAll">Remove All</button>' +
            "</div>" +
            '<fieldset class="win-groupbox vga-gb"><legend>Assigned Vessels</legend>' +
            '<div class="vga-dgv-wrap">' +
            '<table class="win-dgv vga-dgv" id="vgaAsTable">' +
            "<thead><tr>" +
            '<th class="vga-col-id">VesselID</th>' +
            "<th>Vessel Name</th>" +
            "<th>Contents</th>" +
            "</tr></thead>" +
            '<tbody id="vgaAsBody"></tbody>' +
            "</table></div></fieldset>" +
            '<div class="vga-client-footer">' +
            '<button type="button" class="win-btn win-btn-default" id="vgaClose">Close</button>' +
            "</div></div>";

        openAppModal(
            "Assign Vessels to Group - Binventory Workstation",
            html,
            "",
            "modal-vessel-group-assign modal-footer-hidden modal-win-toolwindow"
        );

        var vgaShell = document.querySelector(".vga-shell");
        vgaShell.addEventListener("click", function (e) {
            var trUn = e.target.closest("tr.vga-un-row");
            var trAs = e.target.closest("tr.vga-as-row");
            if (trUn) {
                selectedUnId = trUn.getAttribute("data-vessel-id");
                selectedAsId = null;
                document.querySelectorAll("#vgaUnBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                document.querySelectorAll("#vgaAsBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                trUn.classList.add("selected");
            } else if (trAs) {
                selectedAsId = trAs.getAttribute("data-vessel-id");
                selectedUnId = null;
                document.querySelectorAll("#vgaUnBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                document.querySelectorAll("#vgaAsBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                trAs.classList.add("selected");
            }
        });

        renderVgaTables();

        document.getElementById("vgaAdd").addEventListener("click", function () {
            if (!selectedUnId) {
                alert("Select a Vessel that you wish to add.");
                return;
            }
            working.vesselIds.push(selectedUnId);
            selectedUnId = null;
            renderVgaTables();
        });
        document.getElementById("vgaRemove").addEventListener("click", function () {
            if (!selectedAsId) {
                alert("Select a Vessel that you wish to remove.");
                return;
            }
            working.vesselIds = working.vesselIds.filter(function (id) {
                return id !== selectedAsId;
            });
            selectedAsId = null;
            renderVgaTables();
        });
        document.getElementById("vgaAddAll").addEventListener("click", function () {
            unassignedVessels().forEach(function (v) {
                working.vesselIds.push(v.id);
            });
            renderVgaTables();
        });
        document.getElementById("vgaRemoveAll").addEventListener("click", function () {
            working.vesselIds = [];
            renderVgaTables();
        });
        document.getElementById("vgaClose").addEventListener("click", function () {
            closeAppModal();
            if (typeof onClose === "function") onClose();
        });
    }

    /**
     * frmGroupSetup — Vessel Group Setup (485×253). msAction "1" = new, "2" = edit.
     * restoreParent: scroll/selection to restore on Group Maintenance when Cancel (mirrors VB grid refresh).
     */
    function openGroupSetup(msAction, groupId, restoredWorking, restoreParent) {
        var working = restoredWorking;
        if (!working) {
            if (msAction === "2") {
                var g = findGroupById(groupId);
                if (!g) {
                    closeAppModal();
                    openGroupMaintenance();
                    return;
                }
                working = {
                    isNew: false,
                    id: g.id,
                    name: g.name || "",
                    vesselIds: (g.vesselIds || []).slice(),
                    siteId: g.siteId || state.currentWorkstationSiteId
                };
            } else {
                working = {
                    isNew: true,
                    id: nextGroupRowId(),
                    name: "",
                    vesselIds: [],
                    siteId: state.currentWorkstationSiteId
                };
            }
        }

        var ro = currentUserAccessLevel() === "readonly";

        function renderGsVesselRows() {
            var tbody = document.getElementById("gsVesselBody");
            if (!tbody) return;
            var rows = working.vesselIds
                .map(function (vid) {
                    return findVessel(vid);
                })
                .filter(Boolean);
            if (!rows.length) {
                tbody.innerHTML =
                    '<tr class="gs-empty"><td class="gs-col-id"></td><td colspan="2">No vessels assigned</td></tr>';
                return;
            }
            tbody.innerHTML = rows
                .map(function (v) {
                    return (
                        "<tr data-vessel-id=\"" +
                        escapeHtml(v.id) +
                        "\">" +
                        '<td class="gs-col-id">' +
                        escapeHtml(v.id) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.name) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(v.contents || "") +
                        "</td></tr>"
                    );
                })
                .join("");
        }

        var html =
            '<div class="gs-shell">' +
            '<div class="gs-lbl-title">Vessel Group Setup</div>' +
            '<div class="gs-name-row">' +
            '<label for="gsName">Group Name:</label>' +
            '<input type="text" id="gsName" class="gs-input-name" value="' +
            escapeHtml(working.name) +
            '"' +
            (ro ? " disabled" : "") +
            "></div>" +
            '<fieldset class="win-groupbox gs-gb"><legend>Assigned Vessels</legend>' +
            '<div class="gs-dgv-wrap">' +
            '<table class="win-dgv gs-dgv" id="gsVesselTable">' +
            "<thead><tr>" +
            '<th class="gs-col-id">VesselID</th>' +
            "<th>Vessel Name</th>" +
            "<th>Contents</th>" +
            "</tr></thead>" +
            '<tbody id="gsVesselBody"></tbody>' +
            "</table></div></fieldset>" +
            '<div class="gs-footer-wrap">' +
            '<button type="button" class="win-btn" id="gsAssign"' +
            (ro ? " disabled" : "") +
            ">Assign Vessels</button>" +
            '<div class="gs-footer-right">' +
            '<button type="button" class="win-btn" id="gsSave"' +
            (ro ? " disabled" : "") +
            ">Save</button>" +
            '<button type="button" class="win-btn" id="gsCancel">Cancel</button>' +
            "</div></div></div>";

        openAppModal(
            "Vessel Group Setup - Binventory Workstation",
            html,
            "",
            "modal-group-setup modal-footer-hidden modal-win-toolwindow"
        );

        renderGsVesselRows();

        document.getElementById("gsAssign").addEventListener("click", function () {
            if (ro) return;
            working.name = document.getElementById("gsName").value.trim();
            closeAppModal();
            openVesselGroupAssignDialog(working, function () {
                openGroupSetup(msAction, groupId, working, restoreParent);
            });
        });
        document.getElementById("gsSave").addEventListener("click", function () {
            if (ro) return;
            var nm = document.getElementById("gsName").value.trim();
            if (!nm) {
                alert("Please enter a group name.");
                return;
            }
            working.name = nm;
            var wasNew = working.isNew;
            if (wasNew) {
                state.groups.push({
                    id: working.id,
                    name: working.name,
                    vesselIds: working.vesselIds.slice(),
                    siteId: working.siteId
                });
                working.isNew = false;
            } else {
                var g = findGroupById(working.id);
                if (g) {
                    g.name = working.name;
                    g.vesselIds = working.vesselIds.slice();
                }
            }
            saveState();
            closeAppModal();
            if (wasNew) {
                openGroupMaintenance({ selectGroupId: working.id });
            } else {
                openGroupMaintenance(
                    restoreParent
                        ? Object.assign({}, restoreParent, {
                              selectGroupId: working.id
                          })
                        : { selectGroupId: working.id }
                );
            }
            toast("Group saved (simulation).");
        });
        document.getElementById("gsCancel").addEventListener("click", function () {
            closeAppModal();
            openGroupMaintenance(restoreParent || {});
        });
    }

    /**
     * frmGroupMaintenance — grid GroupID + Group Name; Select / Add New / Delete / Close (584×336).
     */
    function openGroupMaintenance(opts) {
        opts = opts || {};
        var ro = currentUserAccessLevel() === "readonly";

        var list = groupsListSorted();
        var rows = list
            .map(function (g) {
                return (
                    '<tr class="gm-row" data-group-id="' +
                    escapeHtml(g.id) +
                    '">' +
                    '<td class="gm-col-id">' +
                    escapeHtml(String(parseGroupNumericId(g))) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(g.name) +
                    "</td></tr>"
                );
            })
            .join("");
        if (!rows) {
            rows =
                '<tr class="gm-empty"><td class="gm-col-id"></td><td>No groups</td></tr>';
        }

        var html =
            '<div class="sm-win-client">' +
            '<div class="sm-lbl-title">Vessel Group Maintenance</div>' +
            '<div class="sm-main-row">' +
            '<div class="gm-dgv-wrap">' +
            '<table class="win-dgv gm-table" id="gmGroupTable">' +
            "<thead><tr>" +
            '<th class="gm-col-id">GroupID</th>' +
            "<th>Group Name</th>" +
            "</tr></thead>" +
            "<tbody>" +
            rows +
            "</tbody></table></div>" +
            '<div class="sm-actions-col">' +
            '<div class="sm-actions">' +
            '<button type="button" class="win-btn" id="gmBtnSelect" disabled>Select</button>' +
            '<button type="button" class="win-btn" id="gmBtnAdd"' +
            (ro ? " disabled" : "") +
            ">Add New</button>" +
            '<button type="button" class="win-btn" id="gmBtnDelete" disabled>Delete</button>' +
            "</div>" +
            '<button type="button" class="win-btn win-btn-default" id="gmBtnClose" data-close-app>Close</button>' +
            "</div></div></div>";

        openAppModal(
            "Group Maintenance - Binventory Workstation",
            html,
            "",
            "modal-group-maint modal-footer-hidden modal-win-toolwindow"
        );

        var tbody = document.querySelector("#gmGroupTable tbody");
        var wrap = document.querySelector(".gm-dgv-wrap");
        var btnSel = document.getElementById("gmBtnSelect");
        var btnDel = document.getElementById("gmBtnDelete");

        function selectedGroupId() {
            var tr = tbody.querySelector("tr.gm-row.selected");
            return tr ? tr.getAttribute("data-group-id") : null;
        }

        function syncGmButtons() {
            var id = selectedGroupId();
            var on = !!id;
            btnSel.disabled = !on;
            btnDel.disabled = !on || ro;
        }

        function applyGmRestore() {
            if (opts.restoreScrollTop != null && wrap) {
                wrap.scrollTop = opts.restoreScrollTop;
            }
            var gmRows = tbody.querySelectorAll("tr.gm-row");
            if (opts.selectGroupId) {
                gmRows.forEach(function (r) {
                    r.classList.remove("selected");
                    if (r.getAttribute("data-group-id") === opts.selectGroupId) {
                        r.classList.add("selected");
                        r.scrollIntoView({ block: "nearest" });
                    }
                });
            } else if (opts.restoreSelectionIndex >= 0 && gmRows[opts.restoreSelectionIndex]) {
                gmRows.forEach(function (r) {
                    r.classList.remove("selected");
                });
                gmRows[opts.restoreSelectionIndex].classList.add("selected");
            }
            syncGmButtons();
        }

        tbody.addEventListener("click", function (e) {
            var tr = e.target.closest("tr.gm-row");
            if (!tr) return;
            tbody.querySelectorAll("tr.gm-row").forEach(function (r) {
                r.classList.remove("selected");
            });
            tr.classList.add("selected");
            syncGmButtons();
        });

        tbody.addEventListener("dblclick", function (e) {
            var tr = e.target.closest("tr.gm-row");
            if (!tr) return;
            var gid = tr.getAttribute("data-group-id");
            if (!gid) return;
            var idx = Array.prototype.indexOf.call(tbody.querySelectorAll("tr.gm-row"), tr);
            var st = wrap ? wrap.scrollTop : 0;
            openGroupSetup("2", gid, null, {
                restoreSelectionIndex: idx,
                restoreScrollTop: st
            });
        });

        document.getElementById("gmBtnSelect").addEventListener("click", function () {
            var gid = selectedGroupId();
            if (!gid) return;
            var tr = tbody.querySelector("tr.gm-row.selected");
            var idx = tr ? Array.prototype.indexOf.call(tbody.querySelectorAll("tr.gm-row"), tr) : -1;
            var st = wrap ? wrap.scrollTop : 0;
            openGroupSetup("2", gid, null, {
                restoreSelectionIndex: idx,
                restoreScrollTop: st
            });
        });

        document.getElementById("gmBtnAdd").addEventListener("click", function () {
            if (ro) return;
            var trSel = tbody.querySelector("tr.gm-row.selected");
            var idx = trSel
                ? Array.prototype.indexOf.call(tbody.querySelectorAll("tr.gm-row"), trSel)
                : -1;
            var st = wrap ? wrap.scrollTop : 0;
            openGroupSetup("1", null, null, {
                restoreSelectionIndex: idx,
                restoreScrollTop: st
            });
        });

        document.getElementById("gmBtnDelete").addEventListener("click", function () {
            var gid = selectedGroupId();
            if (!gid || ro) return;
            var g = findGroupById(gid);
            var nm = g ? g.name : gid;
            if (
                !confirm(
                    "Are you sure you want to delete the group named '" + nm + "'?"
                )
            ) {
                return;
            }
            var st = wrap ? wrap.scrollTop : 0;
            removeGroupIdFromSchedules(gid);
            state.groups = state.groups.filter(function (x) {
                return x.id !== gid;
            });
            saveState();
            closeAppModal();
            openGroupMaintenance({ restoreScrollTop: st });
        });

        applyGmRestore();
    }

    /** miEventType = 1 (measurement) from frmMain / frmScheduleMaintenance — same list as GetScheduleSelectionData(..., 1). */
    var MEASUREMENT_EVENT_TYPE = 1;
    /** miEventType = 3 — frmEmailReports → Site Status Report → frmScheduleMaintenance (BLL GetScheduleSelectionData ..., 3). */
    var EMAIL_REPORT_EVENT_TYPE = 3;

    function schedulesMeasurementList() {
        return state.schedules.filter(function (s) {
            return (
                (s.eventType == null || s.eventType === MEASUREMENT_EVENT_TYPE) &&
                !s.draft
            );
        });
    }

    function schedulesSiteStatusList() {
        return state.schedules.filter(function (s) {
            return s.eventType === EMAIL_REPORT_EVENT_TYPE && !s.draft;
        });
    }

    function findScheduleById(id) {
        return state.schedules.filter(function (s) {
            return s.id === id;
        })[0];
    }

    /** Groups for current workstation site (mirrors tblGroup.WorkstationSiteID filter). */
    function groupsForWorkstationSite() {
        var sid = state.currentWorkstationSiteId;
        return state.groups.filter(function (g) {
            return g.siteId == null || g.siteId === sid;
        });
    }

    /**
     * frmScheduleGroup — Assign Groups to Measurement Schedule (584×439).
     * onClose: reopen Measurement Schedule Setup (same as VB ShowDialog return).
     */
    function openScheduleGroupAssignDialog(scheduleId, onClose, stack) {
        var useStack = !!stack;
        var sch = findScheduleById(scheduleId);
        if (!sch) {
            if (onClose) onClose();
            return;
        }
        ensureScheduleFields(sch);
        if (!sch.groupIds) sch.groupIds = [];

        var selectedUnId = null;
        var selectedAsId = null;

        function unassignedGroups() {
            var assigned = {};
            sch.groupIds.forEach(function (gid) {
                assigned[gid] = true;
            });
            return groupsForWorkstationSite().filter(function (g) {
                return !assigned[g.id];
            });
        }

        function assignedGroupsOrdered() {
            return sch.groupIds
                .map(function (gid) {
                    return state.groups.filter(function (g) {
                        return g.id === gid;
                    })[0];
                })
                .filter(Boolean);
        }

        function renderSgaTables() {
            var unBody = document.getElementById("sgaUnBody");
            var asBody = document.getElementById("sgaAsBody");
            if (!unBody || !asBody) return;
            unBody.innerHTML = unassignedGroups()
                .map(function (g) {
                    var sel = g.id === selectedUnId ? " selected" : "";
                    return (
                        '<tr class="sga-row sga-un-row' +
                        sel +
                        '" data-group-id="' +
                        escapeHtml(g.id) +
                        '">' +
                        '<td class="sga-col-id">' +
                        escapeHtml(g.id) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(g.name) +
                        "</td></tr>"
                    );
                })
                .join("");
            asBody.innerHTML = assignedGroupsOrdered()
                .map(function (g) {
                    var sel = g.id === selectedAsId ? " selected" : "";
                    return (
                        '<tr class="sga-row sga-as-row' +
                        sel +
                        '" data-group-id="' +
                        escapeHtml(g.id) +
                        '">' +
                        '<td class="sga-col-id">' +
                        escapeHtml(g.id) +
                        "</td>" +
                        "<td>" +
                        escapeHtml(g.name) +
                        "</td></tr>"
                    );
                })
                .join("");
        }

        var html =
            '<div class="sga-shell">' +
            '<div class="sga-lbl-title">Assign Groups to Measurement Schedule</div>' +
            '<fieldset class="win-groupbox sga-gb"><legend>Unassigned Groups</legend>' +
            '<div class="sga-dgv-wrap">' +
            '<table class="win-dgv sga-dgv" id="sgaUnTable">' +
            "<thead><tr>" +
            '<th class="sga-col-id">GroupID</th>' +
            "<th>Group Name</th>" +
            "</tr></thead>" +
            '<tbody id="sgaUnBody"></tbody>' +
            "</table></div></fieldset>" +
            '<div class="sga-btn-row">' +
            '<button type="button" class="win-btn" id="sgaAdd">Add</button>' +
            '<button type="button" class="win-btn" id="sgaRemove">Remove</button>' +
            '<button type="button" class="win-btn" id="sgaAddAll">Add All</button>' +
            '<button type="button" class="win-btn" id="sgaRemoveAll">Remove All</button>' +
            "</div>" +
            '<fieldset class="win-groupbox sga-gb"><legend>Assigned Groups</legend>' +
            '<div class="sga-dgv-wrap">' +
            '<table class="win-dgv sga-dgv" id="sgaAsTable">' +
            "<thead><tr>" +
            '<th class="sga-col-id">GroupID</th>' +
            "<th>Group Name</th>" +
            "</tr></thead>" +
            '<tbody id="sgaAsBody"></tbody>' +
            "</table></div></fieldset>" +
            '<div class="sga-client-footer">' +
            '<button type="button" class="win-btn win-btn-default" id="sgaClose">Close</button>' +
            "</div></div>";

        if (useStack) {
            openStackedAppModal2(
                "Assign Groups to Measurement Schedule - Binventory Workstation",
                html,
                "",
                "modal-schedule-group-assign modal-footer-hidden modal-win-toolwindow"
            );
        } else {
            openStackedAppModal(
                "Assign Groups to Measurement Schedule - Binventory Workstation",
                html,
                "",
                "modal-schedule-group-assign modal-footer-hidden modal-win-toolwindow"
            );
        }

        var sgaBody = useStack ? appModalBodyStack2 : appModalBodyStack;
        var sgaShell = sgaBody.querySelector(".sga-shell");
        sgaShell.addEventListener("click", function (e) {
            var trUn = e.target.closest("tr.sga-un-row");
            var trAs = e.target.closest("tr.sga-as-row");
            if (trUn) {
                selectedUnId = trUn.getAttribute("data-group-id");
                selectedAsId = null;
                document.querySelectorAll("#sgaUnBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                document.querySelectorAll("#sgaAsBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                trUn.classList.add("selected");
            } else if (trAs) {
                selectedAsId = trAs.getAttribute("data-group-id");
                selectedUnId = null;
                document.querySelectorAll("#sgaUnBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                document.querySelectorAll("#sgaAsBody tr").forEach(function (r) {
                    r.classList.remove("selected");
                });
                trAs.classList.add("selected");
            }
        });

        renderSgaTables();

        document.getElementById("sgaAdd").addEventListener("click", function () {
            if (!selectedUnId) {
                alert("Select a Group that you wish to add.");
                return;
            }
            sch.groupIds.push(selectedUnId);
            selectedUnId = null;
            saveState();
            renderSgaTables();
        });
        document.getElementById("sgaRemove").addEventListener("click", function () {
            if (!selectedAsId) {
                alert("Select a Group that you wish to remove.");
                return;
            }
            sch.groupIds = sch.groupIds.filter(function (id) {
                return id !== selectedAsId;
            });
            selectedAsId = null;
            saveState();
            renderSgaTables();
        });
        document.getElementById("sgaAddAll").addEventListener("click", function () {
            unassignedGroups().forEach(function (g) {
                sch.groupIds.push(g.id);
            });
            saveState();
            renderSgaTables();
        });
        document.getElementById("sgaRemoveAll").addEventListener("click", function () {
            sch.groupIds = [];
            saveState();
            renderSgaTables();
        });
        document.getElementById("sgaClose").addEventListener("click", function () {
            saveState();
            closeMssOrAssignLayer(useStack);
            if (typeof onClose === "function") onClose();
        });
    }

    /** Draft schedule row (not shown in Scheduler Maintenance until Save). */
    function createDraftMeasurementSchedule(miEventType, stack) {
        var et = miEventType != null ? miEventType : MEASUREMENT_EVENT_TYPE;
        var draft = {
            id: nextScheduleRowId(),
            draft: true,
            name: "",
            groupIds: [],
            vesselIds: vesselsForCurrentSiteSorted()
                .slice(0, 2)
                .map(function (v) {
                    return v.id;
                }),
            eventType: et,
            scheduleActive: true,
            scheduleType: 1,
            scheduleDays: newScheduleDaysUnchecked(),
            scheduleStartDate: todayIsoDateLocal(),
            scheduleInfoTime: nowTimeHHMM(),
            timeIntervalMinutes: "60",
            timeIntervalStart: "08:00",
            timeIntervalEnd: "17:00",
            time: "08:00"
        };
        state.schedules.push(draft);
        openMeasurementScheduleSetup("1", draft.id, et, stack);
    }

    /** Persists Measurement Schedule Setup fields from DOM into schedule row (for Assign Groups + Save). */
    function syncMeasurementScheduleFromDom(sch) {
        if (!sch || !document.getElementById("mssName")) return;
        sch.name = document.getElementById("mssName").value.trim();
        var tEl = document.querySelector('input[name="mssScheduleType"]:checked');
        sch.scheduleType = tEl ? parseInt(tEl.value, 10) : 1;
        if (isNaN(sch.scheduleType)) sch.scheduleType = 1;

        var dayEls = document.querySelectorAll(".mss-day");
        sch.scheduleDays = [];
        for (var di = 0; di < dayEls.length; di++) {
            sch.scheduleDays.push(!!dayEls[di].checked);
        }
        while (sch.scheduleDays.length < 7) sch.scheduleDays.push(false);

        sch.scheduleStartDate = document.getElementById("mssStartDate").value || todayIsoDateLocal();
        var infoTime = document.getElementById("mssInfoStartTime").value || "08:00";
        sch.scheduleInfoTime = infoTime.length >= 5 ? infoTime.slice(0, 5) : infoTime;
        var actEl = document.getElementById("mssActive");
        sch.scheduleActive = actEl.disabled ? true : actEl.checked;

        var tInt =
            sch.timeIntervalMinutes != null ? String(sch.timeIntervalMinutes) : "60";
        var tStart = sch.timeIntervalStart ? sch.timeIntervalStart.slice(0, 5) : "08:00";
        var tEnd = sch.timeIntervalEnd ? sch.timeIntervalEnd.slice(0, 5) : "17:00";
        if (sch.scheduleType === 1) {
            tInt = document.getElementById("mssIntervalMinutes").value.trim() || "60";
            tStart = document.getElementById("mssIntervalStart").value || "08:00";
            tEnd = document.getElementById("mssIntervalEnd").value || "17:00";
            if (tStart.length >= 5) tStart = tStart.slice(0, 5);
            if (tEnd.length >= 5) tEnd = tEnd.slice(0, 5);
        }
        sch.timeIntervalMinutes = tInt;
        sch.timeIntervalStart = tStart;
        sch.timeIntervalEnd = tEnd;
        sch.time = sch.scheduleType === 1 ? tStart : sch.scheduleInfoTime;
    }

    function nextScheduleRowId() {
        var max = 0;
        state.schedules.forEach(function (s) {
            var n = parseInt(String(s.id || "").replace(/^s/i, ""), 10);
            if (!isNaN(n) && n > max) max = n;
        });
        return "s" + (max + 1);
    }

    function parseScheduleNumericId(sch) {
        var n = parseInt(String(sch.id || "").replace(/^s/i, ""), 10);
        return isNaN(n) ? 0 : n;
    }

    function defaultScheduleDays() {
        return [true, true, true, true, true, true, true];
    }

    function newScheduleDaysUnchecked() {
        return [false, false, false, false, false, false, false];
    }

    function todayIsoDateLocal() {
        var d = new Date();
        var y = d.getFullYear();
        var mo = ("0" + (d.getMonth() + 1)).slice(-2);
        var day = ("0" + d.getDate()).slice(-2);
        return y + "-" + mo + "-" + day;
    }

    function nowTimeHHMM() {
        var d = new Date();
        return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
    }

    /** HH:MM 24h for schedule fields (mirrors DateTimePicker value, not clock UI). */
    function mssNormalizeHHMM(hhmm) {
        var m = String(hhmm || "").trim().match(/^(\d{1,2}):(\d{2})/);
        var h = m ? parseInt(m[1], 10) : 8;
        var mi = m ? parseInt(m[2], 10) : 0;
        if (isNaN(h) || isNaN(mi)) {
            h = 8;
            mi = 0;
        }
        h = ((h % 24) + 24) % 24;
        mi = ((mi % 60) + 60) % 60;
        return ("0" + h).slice(-2) + ":" + ("0" + mi).slice(-2);
    }

    function mssAddMinutesHHMM(hhmm, delta) {
        var n = mssNormalizeHHMM(hhmm).match(/^(\d{2}):(\d{2})$/);
        var total = parseInt(n[1], 10) * 60 + parseInt(n[2], 10) + delta;
        total = ((total % 1440) + 1440) % 1440;
        var nh = Math.floor(total / 60);
        var nm = total % 60;
        return ("0" + nh).slice(-2) + ":" + ("0" + nm).slice(-2);
    }

    /** Fixed-width display so each digit column maps to hour / minute / AM-PM (DateTimePicker ShowUpDown). */
    function mssFormat24To12(hhmm) {
        var n = mssNormalizeHHMM(hhmm).match(/^(\d{2}):(\d{2})$/);
        var h = parseInt(n[1], 10);
        var mi = parseInt(n[2], 10);
        var ampm = h >= 12 ? "PM" : "AM";
        var h12 = h % 12;
        if (h12 === 0) h12 = 12;
        return ("0" + h12).slice(-2) + ":" + ("0" + mi).slice(-2) + " " + ampm;
    }

    function mssValidHourDigits(d0, d1) {
        if (d0 === 0) return d1 >= 1 && d1 <= 9;
        if (d0 === 1) return d1 >= 0 && d1 <= 2;
        return false;
    }

    /** Spin one digit of 12-hour clock (1–12); invalid combo falls back to ±1 hour. */
    function mssSpinHourDigit12(h12, digitIdx, up) {
        var s = ("0" + h12).slice(-2);
        var d0 = parseInt(s.charAt(0), 10);
        var d1 = parseInt(s.charAt(1), 10);
        var dir = up ? 1 : -1;
        if (digitIdx === 0) {
            d0 = (d0 + dir + 2) % 2;
        } else {
            d1 = (d1 + dir + 10) % 10;
        }
        if (mssValidHourDigits(d0, d1)) {
            return d0 * 10 + d1;
        }
        return ((h12 - 1 + dir + 12) % 12) + 1;
    }

    /** Map caret index in "hh:mm AM" string to spin zone: hour digit 0–1, minute 3–4, ampm 6+. */
    function mssNormalizeTimeCaretPos(pos, displayLen) {
        if (pos == null || pos < 0) return 4;
        if (pos >= displayLen) return 6;
        if (pos <= 1) return pos;
        if (pos === 2) return 1;
        if (pos === 3 || pos === 4) return pos;
        if (pos === 5) return 4;
        return 6;
    }

    /**
     * Up/down affects the digit (or AM/PM segment) at the caret — not only the minutes “ones” place.
     * Minute tens: ±10 min; minute ones: ±1 min (with rollover). Hour digits: spin that decimal digit with 12h rules.
     */
    function mssDigitSpinFromCaret(hhmm24, caretPos, up) {
        var n = mssNormalizeHHMM(hhmm24).match(/^(\d{2}):(\d{2})$/);
        var h24 = parseInt(n[1], 10);
        var mi = parseInt(n[2], 10);
        var disp = mssFormat24To12(hhmm24);
        var pos = mssNormalizeTimeCaretPos(caretPos, disp.length);

        if (pos >= 6) {
            var nh24 = h24 < 12 ? h24 + 12 : h24 - 12;
            return ("0" + nh24).slice(-2) + ":" + ("0" + mi).slice(-2);
        }

        var wasPm = h24 >= 12;
        var h12 = h24 % 12;
        if (h12 === 0) h12 = 12;

        if (pos <= 1) {
            var nh12 = mssSpinHourDigit12(h12, pos, up);
            var nh24b;
            if (wasPm) {
                nh24b = nh12 === 12 ? 12 : nh12 + 12;
            } else {
                nh24b = nh12 === 12 ? 0 : nh12;
            }
            return ("0" + nh24b).slice(-2) + ":" + ("0" + mi).slice(-2);
        }

        var delta = pos === 3 ? (up ? 10 : -10) : up ? 1 : -1;
        var total = h24 * 60 + mi + delta;
        total = ((total % 1440) + 1440) % 1440;
        var nh = Math.floor(total / 60);
        var nm = total % 60;
        return ("0" + nh).slice(-2) + ":" + ("0" + nm).slice(-2);
    }

    function mssTimeSpinHtml(id, valueHHMM) {
        var v = escapeHtml(mssNormalizeHHMM(valueHHMM));
        var dispId = id + "Display";
        return (
            '<div class="mss-time-spin">' +
            '<input type="hidden" id="' +
            id +
            '" value="' +
            v +
            '">' +
            '<input type="text" class="mss-time-display" id="' +
            dispId +
            '" readonly tabindex="0" value="">' +
            '<div class="mss-spin-btns">' +
            '<button type="button" class="mss-spin-up" aria-label="Increase time">\u25B2</button>' +
            '<button type="button" class="mss-spin-down" aria-label="Decrease time">\u25BC</button>' +
            "</div></div>"
        );
    }

    function initMssTimeSpinners() {
        document.querySelectorAll(".mss-time-spin").forEach(function (wrap) {
            var hid = wrap.querySelector('input[type="hidden"]');
            var disp = wrap.querySelector(".mss-time-display");
            if (!hid || !disp) return;
            var lastCaret = 4;

            function sync() {
                hid.value = mssNormalizeHHMM(hid.value);
                disp.value = mssFormat24To12(hid.value);
            }

            function captureCaret() {
                if (document.activeElement === disp && typeof disp.selectionStart === "number") {
                    lastCaret = disp.selectionStart;
                }
            }

            disp.addEventListener("mouseup", captureCaret);
            disp.addEventListener("keyup", captureCaret);
            disp.addEventListener("select", captureCaret);

            function spin(up) {
                if (hid.disabled) return;
                var pos =
                    document.activeElement === disp && typeof disp.selectionStart === "number"
                        ? disp.selectionStart
                        : lastCaret;
                hid.value = mssDigitSpinFromCaret(hid.value, pos, up);
                sync();
                var len = disp.value.length;
                var restore = Math.min(pos, len);
                disp.focus();
                try {
                    disp.setSelectionRange(restore, restore);
                } catch (e) {
                    /* ignore */
                }
                lastCaret = restore;
            }

            sync();
            lastCaret = Math.min(lastCaret, disp.value.length);

            var up = wrap.querySelector(".mss-spin-up");
            var dn = wrap.querySelector(".mss-spin-down");
            if (up) {
                up.addEventListener("mousedown", function (e) {
                    e.preventDefault();
                    captureCaret();
                });
                up.addEventListener("click", function () {
                    spin(true);
                });
            }
            if (dn) {
                dn.addEventListener("mousedown", function (e) {
                    e.preventDefault();
                    captureCaret();
                });
                dn.addEventListener("click", function () {
                    spin(false);
                });
            }
        });
    }

    function setMssTimeSpinDisabled(hiddenEl, isDisabled) {
        if (!hiddenEl) return;
        hiddenEl.disabled = !!isDisabled;
        var wrap = hiddenEl.closest(".mss-time-spin");
        if (wrap) wrap.classList.toggle("mss-time-spin-disabled", !!isDisabled);
    }

    function ensureScheduleFields(s) {
        if (!s) return;
        if (s.eventType == null) s.eventType = MEASUREMENT_EVENT_TYPE;
        if (s.scheduleActive == null) s.scheduleActive = true;
        if (s.scheduleType == null) s.scheduleType = 1;
        if (!s.scheduleDays || s.scheduleDays.length !== 7) s.scheduleDays = defaultScheduleDays();
        if (!s.vesselIds) s.vesselIds = [];
        if (s.time == null) s.time = "08:00";
        if (s.scheduleInfoTime == null) s.scheduleInfoTime = s.time || "08:00";
        if (s.scheduleStartDate == null) s.scheduleStartDate = "";
        if (s.timeIntervalMinutes == null) s.timeIntervalMinutes = "60";
        if (s.timeIntervalStart == null) s.timeIntervalStart = "08:00";
        if (s.timeIntervalEnd == null) s.timeIntervalEnd = "17:00";
        if (!s.groupIds) s.groupIds = [];
    }

    /**
     * frmScheduleSetup — Measurement Schedule Setup. msAction "1" = new draft, "2" = edit saved row.
     * miEventType from frmScheduleMaintenance (1 = measurement, 3 = Site Status Reports from frmEmailReports).
     */
    function openMeasurementScheduleSetup(msAction, scheduleId, miEventType, stack) {
        var useStack = !!stack;
        if (!scheduleId) {
            openScheduleMaintenance(
                miEventType != null ? miEventType : MEASUREMENT_EVENT_TYPE,
                { stack: useStack }
            );
            return;
        }
        var sch = findScheduleById(scheduleId);
        if (!sch) {
            openScheduleMaintenance(
                miEventType != null ? miEventType : MEASUREMENT_EVENT_TYPE,
                { stack: useStack }
            );
            return;
        }
        if (msAction === "2" && sch.draft) {
            openScheduleMaintenance(
                miEventType != null ? miEventType : MEASUREMENT_EVENT_TYPE,
                { stack: useStack }
            );
            return;
        }
        ensureScheduleFields(sch);
        var maintainEt =
            miEventType != null
                ? miEventType
                : sch.eventType != null
                  ? sch.eventType
                  : MEASUREMENT_EVENT_TYPE;

        var nameVal = sch.name || "";
        var st = sch.scheduleType != null ? sch.scheduleType : 1;
        var activeChk = sch.scheduleActive !== false;
        var days =
            sch.scheduleDays && sch.scheduleDays.length === 7
                ? sch.scheduleDays.slice()
                : msAction === "1"
                  ? newScheduleDaysUnchecked()
                  : defaultScheduleDays();

        var startDateVal = sch.scheduleStartDate ? sch.scheduleStartDate : todayIsoDateLocal();
        var infoTimeVal = (sch.scheduleInfoTime || sch.time || "08:00").slice(0, 5);
        var intervalMins = sch.timeIntervalMinutes != null ? String(sch.timeIntervalMinutes) : "60";
        var intervalStart = (sch.timeIntervalStart || "08:00").slice(0, 5);
        var intervalEnd = (sch.timeIntervalEnd || "17:00").slice(0, 5);

        var dayLabels = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        var dayChecks = days
            .map(function (on, i) {
                return (
                    '<label class="mss-day-lbl"><input type="checkbox" class="mss-day" data-i="' +
                    i +
                    '"' +
                    (on ? " checked" : "") +
                    "> " +
                    dayLabels[i] +
                    "</label>"
                );
            })
            .join("");

        var types = [
            { v: 1, t: "Regular Intervals" },
            { v: 2, t: "Daily" },
            { v: 3, t: "Weekly" },
            { v: 4, t: "Monthly" },
            { v: 0, t: "Run Once" }
        ];
        var typeRadios = types
            .map(function (x) {
                return (
                    '<label class="mss-type-lbl"><input type="radio" name="mssScheduleType" value="' +
                    x.v +
                    '"' +
                    (st === x.v ? " checked" : "") +
                    "> " +
                    escapeHtml(x.t) +
                    "</label>"
                );
            })
            .join("");

        var html =
            '<div class="mss-form">' +
            '<div class="mss-lbl-title">Measurement Schedule Setup</div>' +
            '<fieldset class="mss-fieldset" id="mssGrpInfo"><legend>Schedule Information</legend>' +
            '<div class="mss-row"><label for="mssName">Schedule Name:</label>' +
            '<input type="text" id="mssName" class="mss-input-name" value="' +
            escapeHtml(nameVal) +
            '"></div>' +
            '<div class="mss-row mss-row-startline">' +
            '<label for="mssStartDate">Start Date:</label>' +
            '<div class="mss-startline-fields">' +
            '<input type="date" id="mssStartDate" value="' +
            escapeHtml(startDateVal) +
            '">' +
            '<label id="mssLblStartTime" class="mss-inline-lbl" for="mssInfoStartTimeDisplay">Start Time:</label>' +
            mssTimeSpinHtml("mssInfoStartTime", infoTimeVal) +
            "</div></div>" +
            '<div class="mss-row mss-row-check"><label></label><label><input type="checkbox" id="mssActive"' +
            (activeChk ? " checked" : "") +
            (sch.draft ? " disabled" : "") +
            "> Schedule Active</label></div>" +
            "</fieldset>" +
            '<fieldset class="mss-fieldset"><legend>Schedule Type</legend>' +
            '<div class="mss-type-row">' +
            typeRadios +
            "</div></fieldset>" +
            '<fieldset class="mss-fieldset" id="mssGrpRunDay"><legend>Run only on</legend>' +
            '<div class="mss-days-row">' +
            dayChecks +
            "</div></fieldset>" +
            '<fieldset class="mss-fieldset" id="mssGrpRunIntervals"><legend>Run in intervals</legend>' +
            '<div class="mss-interval-line">' +
            '<span class="mss-il-part mss-il-intervals">' +
            '<label for="mssIntervalMinutes">Intervals:</label>' +
            '<input type="text" id="mssIntervalMinutes" class="mss-interval-mins" value="' +
            escapeHtml(intervalMins) +
            '">' +
            '<span class="mss-minutes-suffix">minutes</span></span>' +
            '<span class="mss-interval-times">' +
            '<span class="mss-il-part">' +
            '<label class="mss-interval-lbl" for="mssIntervalStartDisplay">Start Time:</label>' +
            mssTimeSpinHtml("mssIntervalStart", intervalStart) +
            "</span>" +
            '<span class="mss-il-part">' +
            '<label class="mss-interval-lbl" for="mssIntervalEndDisplay">End Time:</label>' +
            mssTimeSpinHtml("mssIntervalEnd", intervalEnd) +
            "</span></span></div></fieldset>" +
            "</div>";

        if (useStack) {
            openStackedAppModal2(
                "Measurement Schedule Setup - Binventory Workstation",
                html,
                '<div class="mss-footer-wrap">' +
                    '<button type="button" class="secondary" id="mssAssign">Assign Groups</button>' +
                    '<div class="mss-footer-right">' +
                    '<button type="button" class="primary" id="mssSave">Save</button>' +
                    '<button type="button" class="secondary" id="mssCancel">Cancel</button>' +
                    "</div></div>",
                "modal-measurement-schedule"
            );
        } else {
            openStackedAppModal(
                "Measurement Schedule Setup - Binventory Workstation",
                html,
                '<div class="mss-footer-wrap">' +
                    '<button type="button" class="secondary" id="mssAssign">Assign Groups</button>' +
                    '<div class="mss-footer-right">' +
                    '<button type="button" class="primary" id="mssSave">Save</button>' +
                    '<button type="button" class="secondary" id="mssCancel">Cancel</button>' +
                    "</div></div>",
                "modal-measurement-schedule"
            );
        }

        function applyMssScheduleTypeUI() {
            var tEl = document.querySelector('input[name="mssScheduleType"]:checked');
            var scheduleType = tEl ? parseInt(tEl.value, 10) : 1;
            if (isNaN(scheduleType)) scheduleType = 1;

            var lblInfoStart = document.getElementById("mssLblStartTime");
            var infoStart = document.getElementById("mssInfoStartTime");
            var grpRunDay = document.getElementById("mssGrpRunDay");
            var grpIntervals = document.getElementById("mssGrpRunIntervals");
            var intervalMinsEl = document.getElementById("mssIntervalMinutes");
            var intervalStartEl = document.getElementById("mssIntervalStart");
            var intervalEndEl = document.getElementById("mssIntervalEnd");

            function lblDisabled(on) {
                if (lblInfoStart) lblInfoStart.style.opacity = on ? "0.55" : "";
            }

            switch (scheduleType) {
                case 0:
                    lblDisabled(false);
                    setMssTimeSpinDisabled(infoStart, false);
                    if (grpIntervals) grpIntervals.disabled = true;
                    if (grpRunDay) grpRunDay.disabled = true;
                    break;
                case 1:
                    lblDisabled(true);
                    setMssTimeSpinDisabled(infoStart, true);
                    if (grpIntervals) grpIntervals.disabled = false;
                    if (grpRunDay) grpRunDay.disabled = false;
                    if (intervalMinsEl) intervalMinsEl.disabled = false;
                    if (intervalStartEl) setMssTimeSpinDisabled(intervalStartEl, false);
                    if (intervalEndEl) setMssTimeSpinDisabled(intervalEndEl, false);
                    break;
                case 2:
                    lblDisabled(false);
                    setMssTimeSpinDisabled(infoStart, false);
                    if (grpIntervals) grpIntervals.disabled = true;
                    if (grpRunDay) grpRunDay.disabled = false;
                    break;
                case 3:
                case 4:
                    lblDisabled(false);
                    setMssTimeSpinDisabled(infoStart, false);
                    if (grpIntervals) grpIntervals.disabled = true;
                    if (grpRunDay) grpRunDay.disabled = true;
                    break;
                default:
                    break;
            }
        }

        var radios = document.querySelectorAll('input[name="mssScheduleType"]');
        for (var ri = 0; ri < radios.length; ri++) {
            radios[ri].addEventListener("change", applyMssScheduleTypeUI);
        }
        initMssTimeSpinners();
        applyMssScheduleTypeUI();

        document.getElementById("mssAssign").addEventListener("click", function () {
            syncMeasurementScheduleFromDom(sch);
            saveState();
            var sid = sch.id;
            var action = msAction;
            var etKeep = maintainEt;
            closeMssOrAssignLayer(useStack);
            openScheduleGroupAssignDialog(sid, function () {
                openMeasurementScheduleSetup(action, sid, etKeep, useStack);
            }, useStack);
        });
        document.getElementById("mssCancel").addEventListener("click", function () {
            if (sch.draft) {
                state.schedules = state.schedules.filter(function (s) {
                    return s.id !== sch.id;
                });
            }
            saveState();
            closeMssOrAssignLayer(useStack);
            openScheduleMaintenance(maintainEt, { stack: useStack });
        });
        document.getElementById("mssSave").addEventListener("click", function () {
            var nm = document.getElementById("mssName").value.trim();
            if (!nm) {
                toast("Please enter a schedule name.");
                return;
            }
            var tEl = document.querySelector('input[name="mssScheduleType"]:checked');
            var scheduleType = tEl ? parseInt(tEl.value, 10) : 1;
            if (isNaN(scheduleType)) scheduleType = 1;

            if (scheduleType === 2) {
                var anyDay = false;
                document.querySelectorAll(".mss-day").forEach(function (c) {
                    if (c.checked) anyDay = true;
                });
                if (!anyDay) {
                    toast("Please select the day(s) to run.");
                    return;
                }
            }

            var dayEls = document.querySelectorAll(".mss-day");
            var scheduleDays = [];
            for (var di = 0; di < dayEls.length; di++) {
                scheduleDays.push(!!dayEls[di].checked);
            }
            while (scheduleDays.length < 7) scheduleDays.push(false);

            var startDate = document.getElementById("mssStartDate").value || todayIsoDateLocal();
            var infoTime = document.getElementById("mssInfoStartTime").value || "08:00";
            if (infoTime.length >= 5) infoTime = infoTime.slice(0, 5);
            var actEl = document.getElementById("mssActive");
            var act = actEl.disabled ? true : actEl.checked;

            var tInt =
                sch.timeIntervalMinutes != null ? String(sch.timeIntervalMinutes) : "60";
            var tStart = sch.timeIntervalStart ? sch.timeIntervalStart.slice(0, 5) : "08:00";
            var tEnd = sch.timeIntervalEnd ? sch.timeIntervalEnd.slice(0, 5) : "17:00";
            if (scheduleType === 1) {
                tInt = document.getElementById("mssIntervalMinutes").value.trim() || "60";
                tStart = document.getElementById("mssIntervalStart").value || "08:00";
                tEnd = document.getElementById("mssIntervalEnd").value || "17:00";
                if (tStart.length >= 5) tStart = tStart.slice(0, 5);
                if (tEnd.length >= 5) tEnd = tEnd.slice(0, 5);
            }

            var timeForState = scheduleType === 1 ? tStart : infoTime;

            sch.name = nm;
            sch.eventType = maintainEt;
            sch.time = timeForState;
            sch.scheduleInfoTime = infoTime;
            sch.scheduleStartDate = startDate;
            sch.scheduleActive = act;
            sch.scheduleType = scheduleType;
            sch.scheduleDays = scheduleDays;
            sch.timeIntervalMinutes = tInt;
            sch.timeIntervalStart = tStart;
            sch.timeIntervalEnd = tEnd;
            delete sch.draft;

            saveState();
            closeMssOrAssignLayer(useStack);
            openScheduleMaintenance(maintainEt, { stack: useStack });
            toast("Schedule saved (simulation).");
        });
    }

    /**
     * frmScheduleMaintenance — miEventType 1: "Scheduler Maintenance"; miEventType 3: "Site Status Reports" (frmEmailReports).
     * opts.stack — open on #backdropAppStack so frmEmailReports stays visible underneath (WinForms dialog stack).
     */
    function openScheduleMaintenance(miEventType, opts) {
        var useStack = opts && opts.stack;
        var et = miEventType != null ? miEventType : MEASUREMENT_EVENT_TYPE;
        var msTitle =
            et === EMAIL_REPORT_EVENT_TYPE ? "Site Status Reports" : "Scheduler Maintenance";
        var list =
            et === EMAIL_REPORT_EVENT_TYPE ? schedulesSiteStatusList() : schedulesMeasurementList();
        list.sort(function (a, b) {
            return parseScheduleNumericId(a) - parseScheduleNumericId(b);
        });
        var rows = list
            .map(function (s) {
                ensureScheduleFields(s);
                return (
                    "<tr class=\"sm-sched-row\" data-schedule-id=\"" +
                    escapeHtml(s.id) +
                    "\">" +
                    '<td class="sm-sched-id">' +
                    escapeHtml(s.id) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(s.name) +
                    "</td></tr>"
                );
            })
            .join("");
        if (!rows) {
            rows = '<tr class="sm-sched-empty"><td class="sm-sched-id"></td><td>No schedules</td></tr>';
        }

        var html =
            '<div class="sm-win-client">' +
            '<div class="sm-lbl-title">' +
            escapeHtml(msTitle) +
            "</div>" +
            '<div class="sm-main-row">' +
            '<div class="sm-dgv-wrap">' +
            '<table class="win-dgv sm-sched-table" id="smSchedTable">' +
            "<thead><tr>" +
            '<th class="sm-sched-id">ScheduleID</th>' +
            "<th>Schedule Name</th>" +
            "</tr></thead>" +
            "<tbody>" +
            rows +
            "</tbody></table></div>" +
            '<div class="sm-actions-col">' +
            '<div class="sm-actions">' +
            '<button type="button" class="win-btn" id="smBtnSelect" disabled>Select</button>' +
            '<button type="button" class="win-btn" id="smBtnAdd">Add New</button>' +
            '<button type="button" class="win-btn" id="smBtnDelete" disabled>Delete</button>' +
            "</div>" +
            '<button type="button" class="win-btn win-btn-default" id="smBtnClose" ' +
            (useStack ? "data-close-app-stack" : "data-close-app") +
            ">Close</button>" +
            "</div></div></div>";

        openModalLayer(
            useStack,
            msTitle + " - Binventory Workstation",
            html,
            "",
            "modal-scheduler-maint modal-footer-hidden modal-win-toolwindow"
        );

        var tbody = document.querySelector("#smSchedTable tbody");
        var btnSel = document.getElementById("smBtnSelect");
        var btnDel = document.getElementById("smBtnDelete");

        function selectedId() {
            var tr = tbody.querySelector("tr.sm-sched-row.selected");
            return tr ? tr.getAttribute("data-schedule-id") : null;
        }

        function syncButtons() {
            var id = selectedId();
            var on = !!id;
            btnSel.disabled = !on;
            btnDel.disabled = !on;
        }

        tbody.addEventListener("click", function (e) {
            var tr = e.target.closest("tr.sm-sched-row");
            if (!tr) return;
            tbody.querySelectorAll("tr.sm-sched-row").forEach(function (r) {
                r.classList.remove("selected");
            });
            tr.classList.add("selected");
            syncButtons();
        });

        tbody.addEventListener("dblclick", function (e) {
            var tr = e.target.closest("tr.sm-sched-row");
            if (!tr) return;
            var sid = tr.getAttribute("data-schedule-id");
            if (sid) openMeasurementScheduleSetup("2", sid, et, useStack);
        });

        btnSel.addEventListener("click", function () {
            var sid = selectedId();
            if (sid) openMeasurementScheduleSetup("2", sid, et, useStack);
        });

        document.getElementById("smBtnAdd").addEventListener("click", function () {
            createDraftMeasurementSchedule(et, useStack);
        });

        btnDel.addEventListener("click", function () {
            var sid = selectedId();
            if (!sid) return;
            var sch = state.schedules.filter(function (s) {
                return s.id === sid;
            })[0];
            var nm = sch ? sch.name : sid;
            if (
                !confirm(
                    "Are you sure you want to delete the schedule named '" + nm + "'?"
                )
            ) {
                return;
            }
            state.schedules = state.schedules.filter(function (s) {
                return s.id !== sid;
            });
            saveState();
            closeModalLayer(useStack);
            openScheduleMaintenance(et, { stack: useStack });
        });

        syncButtons();
    }

    /**
     * frmTemporaryGroupSetup — dual listboxes, shuttle buttons, reorder +/−; OK clears gTempGroupVessels and loads selected order (mirrors VesselUtility.ApplyTempGroup).
     */
    function openTemporaryGroup() {
        function allSiteSorted() {
            return vesselsForCurrentSiteSorted();
        }

        var allV = allSiteSorted();
        var selectedIds =
            state.tempGroupVesselIds && state.tempGroupVesselIds.length
                ? state.tempGroupVesselIds.slice()
                : allV.map(function (v) {
                      return v.id;
                  });

        function rebuildLists() {
            var all = allSiteSorted();
            var selSet = {};
            selectedIds.forEach(function (id) {
                selSet[id] = true;
            });
            var avail = all.filter(function (v) {
                return !selSet[v.id];
            });

            var aEl = document.getElementById("tgAvail");
            var sEl = document.getElementById("tgSel");
            while (aEl.firstChild) aEl.removeChild(aEl.firstChild);
            while (sEl.firstChild) sEl.removeChild(sEl.firstChild);
            avail.forEach(function (v) {
                var o = document.createElement("option");
                o.value = v.id;
                o.textContent = v.name || String(v.vesselNumericId || "");
                aEl.appendChild(o);
            });
            selectedIds.forEach(function (id) {
                var v = state.vessels.filter(function (x) {
                    return x.id === id;
                })[0];
                if (!v) return;
                var o = document.createElement("option");
                o.value = v.id;
                o.textContent = v.name || String(v.vesselNumericId || "");
                sEl.appendChild(o);
            });
            enableTgButtons();
        }

        function enableTgButtons() {
            var aEl = document.getElementById("tgAvail");
            var sEl = document.getElementById("tgSel");
            var nAvail = aEl.options.length;
            var nSel = sEl.options.length;
            var ai = aEl.selectedIndex;
            var si = sEl.selectedIndex;

            var bSel = document.getElementById("tgBtnSel");
            var bAll = document.getElementById("tgBtnAll");
            var bRem = document.getElementById("tgBtnRem");
            var bAllRem = document.getElementById("tgBtnAllRem");
            var bUp = document.getElementById("tgBtnUp");
            var bDn = document.getElementById("tgBtnDn");

            if (nAvail > 0) {
                bSel.disabled = ai < 0;
                bAll.disabled = false;
            } else {
                bSel.disabled = true;
                bAll.disabled = true;
            }
            if (nSel > 0) {
                if (si >= 0) {
                    bRem.disabled = false;
                    if (si === 0) {
                        bUp.disabled = true;
                        bDn.disabled = false;
                    } else if (si === nSel - 1) {
                        bUp.disabled = false;
                        bDn.disabled = true;
                    } else {
                        bUp.disabled = false;
                        bDn.disabled = false;
                    }
                } else {
                    bRem.disabled = true;
                    bUp.disabled = true;
                    bDn.disabled = true;
                }
                bAllRem.disabled = false;
            } else {
                bRem.disabled = true;
                bAllRem.disabled = true;
                bUp.disabled = true;
                bDn.disabled = true;
            }
        }

        var html =
            '<div class="tg-setup">' +
            '<div class="tg-lbl-title">Temporary Group Setup</div>' +
            '<div class="tg-lists">' +
            '<div class="tg-list-col">' +
            '<div class="tg-list-cap">Available Vessels</div>' +
            '<select id="tgAvail" class="tg-list" size="8" aria-label="Available Vessels"></select>' +
            "</div>" +
            '<div class="tg-shuttle">' +
            '<button type="button" id="tgBtnSel" class="tg-shuttle-btn" disabled title="Add">&gt;</button>' +
            '<button type="button" id="tgBtnAll" class="tg-shuttle-btn" disabled title="Add all">&gt;&gt;</button>' +
            '<button type="button" id="tgBtnRem" class="tg-shuttle-btn" disabled title="Remove">&lt;</button>' +
            '<button type="button" id="tgBtnAllRem" class="tg-shuttle-btn" disabled title="Remove all">&lt;&lt;</button>' +
            "</div>" +
            '<div class="tg-list-col tg-list-col-sel">' +
            '<div class="tg-list-cap">Selected Vessels</div>' +
            '<select id="tgSel" class="tg-list" size="8" aria-label="Selected Vessels"></select>' +
            "</div>" +
            '<div class="tg-reorder">' +
            '<button type="button" id="tgBtnUp" class="tg-move-btn" disabled title="Move Up">+</button>' +
            '<button type="button" id="tgBtnDn" class="tg-move-btn" disabled title="Move Down">-</button>' +
            "</div></div></div>";

        openAppModal(
            "Temporary Group Setup - Binventory Workstation",
            html,
            '<button type="button" class="primary" id="tgBtnOk">OK</button>' +
                '<button type="button" class="secondary" data-close-app>Cancel</button>',
            "modal-temp-group"
        );

        rebuildLists();

        var tgAvail = document.getElementById("tgAvail");
        var tgSel = document.getElementById("tgSel");

        tgAvail.addEventListener("change", enableTgButtons);
        tgSel.addEventListener("change", enableTgButtons);
        tgAvail.addEventListener("dblclick", function () {
            if (tgAvail.selectedIndex < 0) return;
            var id = tgAvail.options[tgAvail.selectedIndex].value;
            selectedIds.push(id);
            rebuildLists();
            var idx = selectedIds.indexOf(id);
            if (idx >= 0) tgSel.selectedIndex = idx;
        });
        tgSel.addEventListener("dblclick", function () {
            if (tgSel.selectedIndex < 0) return;
            var idx = tgSel.selectedIndex;
            selectedIds.splice(idx, 1);
            rebuildLists();
        });

        document.getElementById("tgBtnSel").addEventListener("click", function () {
            if (tgAvail.selectedIndex < 0) return;
            var id = tgAvail.options[tgAvail.selectedIndex].value;
            selectedIds.push(id);
            rebuildLists();
        });
        document.getElementById("tgBtnAll").addEventListener("click", function () {
            allSiteSorted().forEach(function (v) {
                if (selectedIds.indexOf(v.id) < 0) selectedIds.push(v.id);
            });
            rebuildLists();
        });
        document.getElementById("tgBtnRem").addEventListener("click", function () {
            if (tgSel.selectedIndex < 0) return;
            selectedIds.splice(tgSel.selectedIndex, 1);
            rebuildLists();
        });
        document.getElementById("tgBtnAllRem").addEventListener("click", function () {
            selectedIds = [];
            rebuildLists();
        });
        document.getElementById("tgBtnUp").addEventListener("click", function () {
            var i = tgSel.selectedIndex;
            if (i <= 0) return;
            var t = selectedIds[i - 1];
            selectedIds[i - 1] = selectedIds[i];
            selectedIds[i] = t;
            rebuildLists();
            tgSel.selectedIndex = i - 1;
            enableTgButtons();
        });
        document.getElementById("tgBtnDn").addEventListener("click", function () {
            var i = tgSel.selectedIndex;
            if (i < 0 || i >= selectedIds.length - 1) return;
            var t = selectedIds[i + 1];
            selectedIds[i + 1] = selectedIds[i];
            selectedIds[i] = t;
            rebuildLists();
            tgSel.selectedIndex = i + 1;
            enableTgButtons();
        });

        document.getElementById("tgBtnOk").addEventListener("click", function () {
            state.tempGroupVesselIds = selectedIds.length ? selectedIds.slice() : null;
            saveState();
            closeAppModal();
            state.currentPage = 0;
            refreshUI();
            toast("Temporary group applied to the vessel display (simulation).");
        });
    }

    function openContactMaintenance() {
        var rows = state.contacts.map(function (c) {
            return "<tr><td>" + escapeHtml(c.name) + "</td><td>" + escapeHtml(c.email) + "</td><td>" + escapeHtml(c.phone) + "</td></tr>";
        }).join("");
        var html =
            '<div class="toolbar"><button type="button" id="cAdd" class="primary">Add contact</button></div>' +
            '<table class="data-table"><thead><tr><th>Name</th><th>Email</th><th>Phone</th></tr></thead><tbody>' +
            rows +
            "</tbody></table>";
        openAppModal("Contact Maintenance", html, '<button type="button" class="secondary" data-close-app>Close</button>');
        document.getElementById("cAdd").addEventListener("click", function () {
            state.contacts.push({ id: uid("c"), name: "New contact", email: "new@example.com", phone: "" });
            saveState();
            closeAppModal();
            openContactMaintenance();
        });
    }

    function countSensorNetworksForSite(siteId) {
        var n = 0;
        state.sensorNetworks.forEach(function (net) {
            ensureNetworkFields(net);
            if (net.siteId === siteId) n++;
        });
        return n;
    }

    function nextWorkstationSiteNumericId() {
        var max = 0;
        state.sites.forEach(function (s) {
            ensureSiteFields(s);
            if (s.workstationSiteId > max) max = s.workstationSiteId;
        });
        return max + 1;
    }

    function openSiteSetupModal(isNew, siteId) {
        var site = isNew ? null : findSite(siteId);
        if (!isNew && !site) {
            openSiteMaintenance();
            return;
        }
        if (site) ensureSiteFields(site);
        var draftId = isNew ? uid("st") : siteId;
        var draftWsId = isNew ? nextWorkstationSiteNumericId() : site.workstationSiteId;

        var nameVal = isNew ? "" : site.name;
        var distId = isNew ? "1" : String(site.distanceUnitsId || "1");
        var companyVal = isNew ? "" : site.companyName;
        var street1 = isNew ? "" : site.streetAddress;
        var street2 = isNew ? "" : site.streetAddress2;
        var cityVal = isNew ? "" : site.city;
        var stateVal = isNew ? "" : site.state;
        var zipVal = isNew ? "" : site.zip;
        var countryVal = isNew ? "" : site.country;
        var ipVal = isNew ? "127.0.0.1" : site.serviceHostIp;
        var portVal = isNew ? "8093" : (site.serviceHostPort || "8093");

        var distOpts = buildDistanceUnitsOptions(distId);

        var html =
            '<div class="su-form">' +
            '<div class="ss-form-title">Site Setup</div>' +
            '<fieldset class="ss-fieldset ss-fieldset-site">' +
            "<legend>Site Details</legend>" +
            '<div class="su-row">' +
            '<label for="ssuName">Name/Description:</label>' +
            '<input type="text" id="ssuName" class="ss-input-full" value="' +
            escapeHtml(nameVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuDist">Distance Units:</label>' +
            '<select id="ssuDist" class="su-select-units">' +
            distOpts +
            "</select>" +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="ss-fieldset ss-fieldset-site">' +
            "<legend>Company &amp; Location</legend>" +
            '<div class="su-row">' +
            '<label for="ssuCompany">Company Name:</label>' +
            '<input type="text" id="ssuCompany" class="ss-input-full" value="' +
            escapeHtml(companyVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuStreet1">Street Address:</label>' +
            '<input type="text" id="ssuStreet1" class="ss-input-full" value="' +
            escapeHtml(street1) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuStreet2">Street Address 2:</label>' +
            '<input type="text" id="ssuStreet2" class="ss-input-full" value="' +
            escapeHtml(street2) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuCity">City/Town:</label>' +
            '<input type="text" id="ssuCity" class="ss-input-full" value="' +
            escapeHtml(cityVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuState">State/Province/Region:</label>' +
            '<input type="text" id="ssuState" class="ss-input-full" value="' +
            escapeHtml(stateVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuZip">Zip/Postal Code:</label>' +
            '<input type="text" id="ssuZip" class="ss-input-zip" value="' +
            escapeHtml(zipVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuCountry">Country/Nation:</label>' +
            '<input type="text" id="ssuCountry" class="ss-input-full" value="' +
            escapeHtml(countryVal) +
            '">' +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="ss-fieldset ss-fieldset-site">' +
            "<legend>Engine Service</legend>" +
            '<div class="su-row">' +
            '<label for="ssuHostIp">Host IP Address:</label>' +
            '<input type="text" id="ssuHostIp" class="ss-input-full" value="' +
            escapeHtml(ipVal) +
            '">' +
            "</div>" +
            '<div class="su-row">' +
            '<label for="ssuHostPort">Host Port:</label>' +
            '<input type="text" id="ssuHostPort" class="ss-input-zip" value="' +
            escapeHtml(portVal) +
            '">' +
            "</div>" +
            "</fieldset>" +
            "</div>";

        var footer =
            '<button type="button" id="ssuSave" class="primary ss-footer-btn"><u>S</u>ave</button>' +
            '<button type="button" id="ssuCancel" class="secondary ss-footer-btn"><u>C</u>ancel</button>';

        openAppModal("Site Setup - Binventory Workstation", html, footer, "modal-site-setup");

        document.getElementById("ssuSave").addEventListener("click", function () {
            var name = document.getElementById("ssuName").value.trim() || "New site";
            var dist = document.getElementById("ssuDist").value;
            var company = document.getElementById("ssuCompany").value.trim();
            var street1v = document.getElementById("ssuStreet1").value.trim();
            var street2v = document.getElementById("ssuStreet2").value.trim();
            var cityv = document.getElementById("ssuCity").value.trim();
            var statev = document.getElementById("ssuState").value.trim();
            var zipv = document.getElementById("ssuZip").value.trim();
            var countryv = document.getElementById("ssuCountry").value.trim();
            var ip = document.getElementById("ssuHostIp").value.trim();
            var port = document.getElementById("ssuHostPort").value.trim() || "8093";

            if (isNew) {
                state.sites.push({
                    id: draftId,
                    name: name,
                    workstationSiteId: draftWsId,
                    serviceHostIp: ip,
                    serviceHostPort: port,
                    companyName: company,
                    streetAddress: street1v,
                    streetAddress2: street2v,
                    city: cityv,
                    state: statev,
                    zip: zipv,
                    country: countryv,
                    distanceUnitsId: dist
                });
                ensureSiteFields(state.sites[state.sites.length - 1]);
                saveState();
                refreshSiteAssignmentMenu();
            } else {
                var s = findSite(siteId);
                if (!s) return;
                s.name = name;
                s.serviceHostIp = ip;
                s.serviceHostPort = port;
                s.companyName = company;
                s.streetAddress = street1v;
                s.streetAddress2 = street2v;
                s.city = cityv;
                s.state = statev;
                s.zip = zipv;
                s.country = countryv;
                s.distanceUnitsId = dist;
                ensureSiteFields(s);
                saveState();
                if (state.currentWorkstationSiteId === siteId) {
                    mergeSystemSettingsFromSite(siteId);
                }
                refreshSiteAssignmentMenu();
            }
            openSiteMaintenance();
        });
        document.getElementById("ssuCancel").addEventListener("click", function () {
            openSiteMaintenance();
        });
    }

    function openSiteMaintenance() {
        state.sites.forEach(ensureSiteFields);
        saveState();

        smSelectedSiteId = null;

        var rows = state.sites
            .map(function (s) {
                ensureSiteFields(s);
                var sel = s.id === smSelectedSiteId ? ' class="sm-selected"' : "";
                return (
                    '<tr data-sid="' +
                    escapeHtml(s.id) +
                    '"' +
                    sel +
                    ">" +
                    "<td>" +
                    escapeHtml(s.name) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(s.serviceHostIp) +
                    "</td>" +
                    "</tr>"
                );
            })
            .join("");

        var html =
            '<div class="sm-win">' +
            '<div class="sm-title">Site Maintenance</div>' +
            '<div class="sm-layout">' +
            '<div class="sm-table-wrap">' +
            '<table class="sm-table" role="grid" aria-label="Sites">' +
            "<thead><tr><th>Name/Description</th><th>Host IP Address</th></tr></thead>" +
            "<tbody>" +
            (rows || '<tr><td colspan="2">&nbsp;</td></tr>') +
            "</tbody>" +
            "</table>" +
            "</div>" +
            '<div class="sm-actions">' +
            '<button type="button" id="smBtnSelect" disabled><u>S</u>elect</button>' +
            '<button type="button" id="smAddNew"><u>A</u>dd New</button>' +
            '<button type="button" id="smBtnDelete" disabled><u>D</u>elete</button>' +
            '<button type="button" class="sm-close" data-close-app><u>C</u>lose</button>' +
            "</div>" +
            "</div>" +
            "</div>";

        openAppModal("Site Maintenance - Binventory Workstation", html, "", "modal-site-maintenance");

        function setSmActionButtons(enabled) {
            var sel = document.getElementById("smBtnSelect");
            if (sel) sel.disabled = !enabled;
            var del = document.getElementById("smBtnDelete");
            if (del) del.disabled = !enabled;
        }

        appModalBody.querySelectorAll(".sm-table tbody tr[data-sid]").forEach(function (tr) {
            tr.addEventListener("click", function (e) {
                e.stopPropagation();
                smSelectedSiteId = tr.getAttribute("data-sid");
                appModalBody.querySelectorAll(".sm-table tbody tr[data-sid]").forEach(function (r) {
                    r.classList.remove("sm-selected");
                });
                tr.classList.add("sm-selected");
                setSmActionButtons(true);
            });
            tr.addEventListener("dblclick", function (e) {
                e.preventDefault();
                e.stopPropagation();
                smSelectedSiteId = tr.getAttribute("data-sid");
                openSiteSetupModal(false, smSelectedSiteId);
            });
        });

        var btnSel = document.getElementById("smBtnSelect");
        if (btnSel) {
            btnSel.addEventListener("click", function () {
                if (smSelectedSiteId) openSiteSetupModal(false, smSelectedSiteId);
            });
        }
        var btnDel = document.getElementById("smBtnDelete");
        if (btnDel) {
            btnDel.addEventListener("click", function () {
                if (!smSelectedSiteId) return;
                if (state.sites.length <= 1) {
                    toast("Keep at least one site.");
                    return;
                }
                if (smSelectedSiteId === state.currentWorkstationSiteId) {
                    toast(
                        "You cannot delete this site because you are currently connected to it. Change the workstation site under System Setup before deleting this site."
                    );
                    return;
                }
                var cnt = countSensorNetworksForSite(smSelectedSiteId);
                var victim = findSite(smSelectedSiteId);
                var siteName = victim ? victim.name : "";
                if (cnt === 1) {
                    toast(
                        "There is one sensor network configured at '" +
                            siteName +
                            "'. You must delete it before this site can be deleted."
                    );
                    return;
                }
                if (cnt > 1) {
                    toast(
                        "There are " +
                            cnt +
                            " sensor networks configured at '" +
                            siteName +
                            "'. You must delete them before this site can be deleted."
                    );
                    return;
                }
                if (!confirm("Are you sure you want to delete '" + (siteName || "this site") + "'?")) {
                    return;
                }
                var delId = smSelectedSiteId;
                state.sites = state.sites.filter(function (x) {
                    return x.id !== delId;
                });
                smSelectedSiteId = null;
                saveState();
                closeAppModal();
                openSiteMaintenance();
                toast("Site deleted (simulation).");
            });
        }

        document.getElementById("smAddNew").addEventListener("click", function () {
            openSiteSetupModal(true);
        });
    }

    /** Order matches Binventory Sensor Network Setup protocol dropdown. */
    var SENSOR_PROTOCOLS = [
        "Protocol B",
        "Protocol A",
        "Modbus/RTU",
        "SPL-100 Push",
        "SPL-200 Push",
        "HART Protocol"
    ];

    var BAUD_VALUES = [1200, 2400, 4800, 9600, 19200, 115200];

    function formatBaudLabel(bps) {
        return Number(bps).toLocaleString("en-US") + " bps";
    }

    function buildProtocolOptionsHtml(selected) {
        var sel = selected || "Protocol A";
        var seen = {};
        var parts = [];
        SENSOR_PROTOCOLS.forEach(function (p) {
            seen[p] = true;
            parts.push(
                "<option value=\"" +
                    escapeHtml(p) +
                    "\"" +
                    (p === sel ? " selected" : "") +
                    ">" +
                    escapeHtml(p) +
                    "</option>"
            );
        });
        if (sel && !seen[sel]) {
            parts.unshift(
                "<option value=\"" + escapeHtml(sel) + "\" selected>" + escapeHtml(sel) + "</option>"
            );
        }
        return parts.join("");
    }

    function buildBaudOptionsHtml(baud) {
        var b = parseInt(baud, 10);
        if (isNaN(b)) b = 2400;
        var inList = false;
        var parts = [];
        BAUD_VALUES.forEach(function (v) {
            if (v === b) inList = true;
            parts.push(
                "<option value=\"" +
                    v +
                    "\"" +
                    (v === b ? " selected" : "") +
                    ">" +
                    formatBaudLabel(v) +
                    "</option>"
            );
        });
        if (!inList) {
            parts.push(
                "<option value=\"" +
                    b +
                    "\" selected>" +
                    formatBaudLabel(b) +
                    "</option>"
            );
        }
        return parts.join("");
    }

    function buildDataBitsOptionsHtml(bits) {
        var d = parseInt(bits, 10);
        if (d !== 7 && d !== 8) d = 8;
        return (
            "<option value=\"7\"" +
            (d === 7 ? " selected" : "") +
            ">7</option>" +
            "<option value=\"8\"" +
            (d === 8 ? " selected" : "") +
            ">8</option>"
        );
    }

    function buildParityOptionsHtml(parity) {
        var p = parity || "None";
        return (
            "<option value=\"None\"" +
            (p === "None" ? " selected" : "") +
            ">None</option>" +
            "<option value=\"Even\"" +
            (p === "Even" ? " selected" : "") +
            ">Even</option>" +
            "<option value=\"Odd\"" +
            (p === "Odd" ? " selected" : "") +
            ">Odd</option>"
        );
    }

    function buildStopBitsOptionsHtml(sb) {
        var s = parseInt(sb, 10);
        if (s !== 1 && s !== 2) s = 1;
        return (
            "<option value=\"1\"" +
            (s === 1 ? " selected" : "") +
            ">1</option>" +
            "<option value=\"2\"" +
            (s === 2 ? " selected" : "") +
            ">2</option>"
        );
    }

    function parityToChar(p) {
        if (p === "Even") return "E";
        if (p === "Odd") return "O";
        return "N";
    }

    function charToParity(c) {
        var u = String(c || "N").trim().toUpperCase();
        if (u === "E") return "Even";
        if (u === "O") return "Odd";
        return "None";
    }

    function parseCommParamsToFields(n) {
        if (typeof n.baud === "number" && typeof n.dataBits === "number") return;
        var parts = String(n.commParams || "2400,8,N,1").split(",");
        n.baud = parseInt(parts[0], 10) || 2400;
        n.dataBits = parseInt(parts[1], 10) || 8;
        if (n.dataBits !== 7 && n.dataBits !== 8) n.dataBits = 8;
        n.parity = charToParity(parts[2]);
        n.stopBits = parseInt(parts[3], 10) || 1;
        if (n.stopBits !== 1 && n.stopBits !== 2) n.stopBits = 1;
    }

    function syncCommParamsString(n) {
        n.commParams = n.baud + "," + n.dataBits + "," + parityToChar(n.parity) + "," + n.stopBits;
    }

    function ensureNetworkFields(n) {
        if (!n.protocol) n.protocol = "Protocol A";
        if (!n.interface) n.interface = "COM1";
        if (!n.commParams) n.commParams = "2400,8,N,1";
        if (!n.name) n.name = "COM1 Sensor Network";
        if (n.requestAttempts == null) n.requestAttempts = 3;
        if (n.responseTimeout == null) n.responseTimeout = 1.5;
        if (n.transmitDelay == null) {
            n.transmitDelay = n.statusRequestDelay != null ? n.statusRequestDelay : 20;
        }
        if (!n.interfaceMode) n.interfaceMode = "local";
        if (n.remoteIp == null) n.remoteIp = "";
        if (n.remotePort == null) n.remotePort = "";
        if (n.dropTimeout == null) n.dropTimeout = 120;
        if (n.disableStuckTop == null) n.disableStuckTop = false;
        if (n.statusRequestDelay == null) n.statusRequestDelay = 1;
        parseCommParamsToFields(n);
        if (n.siteId == null && state.sites && state.sites.length) {
            n.siteId = state.sites[0].id;
        }
    }

    /** Interface column like eBob DB (COM or IP:port for remote serial server). */
    function networkInterfaceDisplayForTable(n) {
        ensureNetworkFields(n);
        if (n.interfaceMode === "remote") {
            var ip = (n.remoteIp || "").trim();
            var port = (n.remotePort || "").trim();
            return ip ? ip + (port ? ":" + port : "") : "";
        }
        return n.interface || "";
    }

    var snSelectedNetworkId = null;
    var snSetupEditingId = null;
    var smSelectedSiteId = null;

    function findNetwork(id) {
        for (var i = 0; i < state.sensorNetworks.length; i++) {
            if (state.sensorNetworks[i].id === id) return state.sensorNetworks[i];
        }
        return null;
    }

    function findSite(id) {
        for (var i = 0; i < state.sites.length; i++) {
            if (state.sites[i].id === id) return state.sites[i];
        }
        return null;
    }

    /** Mirrors DAL.GetDistanceUnits — simulated list for Site Setup combo. */
    var SITE_DISTANCE_UNITS = [
        { id: "1", name: "Inches" },
        { id: "2", name: "Feet" },
        { id: "3", name: "Meters" },
        { id: "4", name: "Centimeters" }
    ];

    function buildDistanceUnitsOptions(selectedId) {
        return SITE_DISTANCE_UNITS.map(function (u) {
            var sel = String(selectedId) === u.id ? " selected" : "";
            return '<option value="' + escapeHtml(u.id) + '"' + sel + ">" + escapeHtml(u.name) + "</option>";
        }).join("");
    }

    function mergeSystemSettingsFromSite(siteId) {
        var s = findSite(siteId);
        if (!s) return;
        ensureSiteFields(s);
        state.systemSettings.companyName = s.companyName || "";
        state.systemSettings.streetAddress = s.streetAddress || "";
        state.systemSettings.streetAddress2 = s.streetAddress2 || "";
        state.systemSettings.city = s.city || "";
        state.systemSettings.state = s.state || "";
        state.systemSettings.zipCode = s.zip || "";
        state.systemSettings.country = s.country || "";
    }

    /** Vessel layout for Warehouse site — distinct from Main Plant. */
    var WAREHOUSE_VESSEL_NAMES = [
        "BEET",
        "BEET East",
        "Sugar Bin 1",
        "Sugar Bin 2",
        "Grain A",
        "Grain B",
        "Pellet Silo",
        "Bulk Storage"
    ];

    /** Seeded demo: one device type per sensor network (vessels match the network, not the full switchable list). */
    var SEED_PROTOCOL_A_SENSOR_TYPE_ID = 1;
    var SEED_MODBUS_RADAR_SENSOR_TYPE_ID = 11;

    function seedWorkspaceMainPlant() {
        state.vessels = [];
        for (var i = 0; i < 32; i++) {
            var v = seedVessel(i);
            v.siteId = "st1";
            v.sensorTypeId = SEED_PROTOCOL_A_SENSOR_TYPE_ID;
            state.vessels.push(v);
        }
        state.groups = [
            { id: "g1", name: "Mill Line A", vesselIds: ["v1", "v2", "v3"], siteId: "st1" }
        ];
        state.schedules = [
            {
                id: "s1",
                name: "Morning sweep",
                time: "08:00",
                vesselIds: ["v1", "v2"],
                eventType: 1,
                scheduleActive: true,
                scheduleType: 1,
                scheduleStartDate: todayIsoDateLocal(),
                scheduleInfoTime: "08:00",
                timeIntervalMinutes: "60",
                timeIntervalStart: "08:00",
                timeIntervalEnd: "17:00",
                scheduleDays: [true, true, true, true, true, false, false],
                groupIds: ["g1"]
            }
        ];
        state.sensorNetworks = [
            {
                id: "n1",
                name: "COM1 Sensor Network",
                protocol: "Protocol A",
                interface: "COM1",
                commParams: "2400,8,N,1",
                siteId: "st1"
            }
        ];
    }

    function seedWorkspaceWarehouse() {
        state.vessels = [];
        for (var i = 0; i < 8; i++) {
            var v = seedVessel(i);
            v.name = WAREHOUSE_VESSEL_NAMES[i] || "Tank " + (i + 1);
            v.product = i < 2 ? "BEET Pulp" : PRODUCTS[(i + 3) % PRODUCTS.length];
            v.contents = v.product;
            v.siteId = "st2";
            v.sensorTypeId = SEED_MODBUS_RADAR_SENSOR_TYPE_ID;
            state.vessels.push(v);
        }
        state.groups = [
            { id: "g1", name: "Dock Line", vesselIds: ["v1", "v2", "v3"], siteId: "st2" }
        ];
        state.schedules = [
            {
                id: "s1",
                name: "Evening sweep",
                time: "17:00",
                vesselIds: ["v1", "v2"],
                eventType: 1,
                scheduleActive: true,
                scheduleType: 1,
                scheduleStartDate: todayIsoDateLocal(),
                scheduleInfoTime: "18:00",
                timeIntervalMinutes: "60",
                timeIntervalStart: "08:00",
                timeIntervalEnd: "17:00",
                scheduleDays: [true, true, true, true, true, false, false],
                groupIds: ["g1"]
            }
        ];
        state.sensorNetworks = [
            {
                id: "n1",
                name: "Plant Ethernet",
                protocol: "Modbus/RTU",
                interface: "10.0.0.50",
                commParams: "9600,8,N,1",
                interfaceMode: "remote",
                remoteIp: "10.0.0.50",
                remotePort: "502",
                siteId: "st2"
            }
        ];
    }

    function seedWorkspaceGeneric(siteId) {
        state.vessels = [];
        for (var i = 0; i < 6; i++) {
            var v = seedVessel(i);
            v.name = "Site Vessel " + (i + 1);
            v.siteId = siteId;
            v.sensorTypeId = SEED_PROTOCOL_A_SENSOR_TYPE_ID;
            state.vessels.push(v);
        }
        state.groups = [
            { id: "g1", name: "Group 1", vesselIds: ["v1", "v2"], siteId: siteId }
        ];
        state.schedules = [
            {
                id: "s1",
                name: "Daily",
                time: "09:00",
                vesselIds: ["v1"],
                eventType: 1,
                scheduleActive: true,
                scheduleType: 1,
                scheduleStartDate: todayIsoDateLocal(),
                scheduleInfoTime: "09:00",
                timeIntervalMinutes: "60",
                timeIntervalStart: "08:00",
                timeIntervalEnd: "17:00",
                scheduleDays: [true, true, true, true, true, false, false],
                groupIds: ["g1"]
            }
        ];
        state.sensorNetworks = [
            {
                id: "n1",
                name: "Default Network",
                protocol: "Protocol A",
                interface: "COM1",
                commParams: "2400,8,N,1",
                siteId: siteId
            }
        ];
    }

    function applySiteWorkspaceTemplate(siteId) {
        if (siteId === "st1") {
            seedWorkspaceMainPlant();
        } else if (siteId === "st2") {
            seedWorkspaceWarehouse();
        } else {
            seedWorkspaceGeneric(siteId);
        }
    }

    /**
     * Simulates Application.Restart + DB load for the selected workstation site (frmMain SiteAssignment).
     * opts.skipRefreshUI — set true during resetStateToFactoryDefaults so the dashboard is not painted until
     * runBinventoryStartupSequence finishes (otherwise all silos render Retracted before "Loading..." stagger).
     */
    function reloadWorkstationForSite(siteId, opts) {
        opts = opts || {};
        var s = findSite(siteId);
        if (!s) return;
        ensureSiteFields(s);
        state.currentWorkstationSiteId = siteId;
        state.tempGroupVesselIds = null;
        mergeSystemSettingsFromSite(siteId);
        applySiteWorkspaceTemplate(siteId);
        state.vessels.forEach(function (v, idx) {
            ensureVesselFields(v, idx);
        });
        state.sensorNetworks.forEach(ensureNetworkFields);
        state.currentPage = 0;
        saveState();
        updateTitleBar();
        if (opts.skipRefreshUI) {
            refreshSiteAssignmentMenu();
            updateMeasureMenuDisabled();
            var grid = document.getElementById("vesselGrid");
            if (grid) grid.innerHTML = "";
            var strip = document.getElementById("tabStrip");
            if (strip) strip.innerHTML = "";
        } else {
            refreshUI();
        }
    }

    /** ToolStrip-style check menu items (not a combo box). */
    function refreshSiteAssignmentMenu() {
        var container = document.getElementById("siteAssignmentMenuItems");
        if (!container) return;
        var cur = state.currentWorkstationSiteId;
        state.sites.forEach(ensureSiteFields);
        container.innerHTML = state.sites
            .map(function (s) {
                var isOn = s.id === cur;
                return (
                    '<button type="button" role="menuitemradio" aria-checked="' +
                    isOn +
                    '" class="sa-menu-item' +
                    (isOn ? " is-active" : "") +
                    '" data-site-id="' +
                    escapeHtml(s.id) +
                    '">' +
                    '<span class="sa-check" aria-hidden="true">' +
                    (isOn ? "\u2713" : "") +
                    "</span>" +
                    '<span class="sa-menu-text">' +
                    escapeHtml(s.name) +
                    "</span>" +
                    "</button>"
                );
            })
            .join("");
    }

    function closeSensorNetworkSetup() {
        document.getElementById("backdropSnSetup").classList.remove("show");
        document.getElementById("snSetupShell").classList.remove("sns-compact");
        snSetupEditingId = null;
    }

    /** Mirrors frmSensorNetworkSetup.vb cboProtocol_SelectedIndexChanged defaults when user changes protocol. */
    function applySensorNetworkProtocolDefaults(p) {
        var baud = document.getElementById("snsBaud");
        var db = document.getElementById("snsDataBits");
        var par = document.getElementById("snsParity");
        var stb = document.getElementById("snsStopBits");
        var att = document.getElementById("snsAtt");
        var rto = document.getElementById("snsRto");
        var std = document.getElementById("snsStatusDelay");
        var drop = document.getElementById("snsDropTimeout");
        var txd = document.getElementById("snsTransmitDelay");
        var stk = document.getElementById("snsDisableStuckTop");
        function setSel(el, val) {
            if (el) el.value = String(val);
        }
        switch (p) {
            case "Protocol B":
                setSel(baud, 2400);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = "3";
                if (rto) rto.value = "2";
                if (drop) drop.value = "120";
                if (stk) stk.checked = false;
                break;
            case "Protocol A":
                setSel(baud, 2400);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = "3";
                if (rto) rto.value = "1.5";
                if (std) std.value = "1";
                break;
            case "Modbus/RTU":
                setSel(baud, 9600);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = "3";
                if (rto) rto.value = "3.5";
                if (txd) txd.value = "20";
                break;
            case "SPL-100 Push":
            case "SPL-200 Push":
                setSel(baud, 9600);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = "3";
                if (rto) rto.value = "1";
                break;
            case "HART Protocol":
                setSel(baud, 1200);
                setSel(db, 8);
                setSel(par, "Odd");
                setSel(stb, 1);
                if (att) att.value = "3";
                if (rto) rto.value = "2";
                if (txd) txd.value = "60";
                break;
            default:
                break;
        }
    }

    /**
     * Mirrors frmSensorNetworkSetup.vb: interface radios + protocol rules + comm group enabled only for local serial.
     */
    function syncSensorNetworkSetupUi(fromProtocolChange) {
        var protoEl = document.getElementById("snsProtocol");
        var p = protoEl ? protoEl.value : "Protocol A";
        var localRad = document.getElementById("snsIfLocal");
        var remoteRad = document.getElementById("snsIfRemote");
        if (!localRad || !remoteRad) return;

        if (p === "Protocol B") {
            localRad.checked = true;
            remoteRad.disabled = true;
        } else {
            remoteRad.disabled = false;
        }

        var local = localRad.checked;

        if (fromProtocolChange) {
            applySensorNetworkProtocolDefaults(p);
        }

        var comEl = document.getElementById("snsCom");
        var ipEl = document.getElementById("snsIp");
        var portEl = document.getElementById("snsPort");
        var portLink = document.getElementById("snsPortLink");
        var lblIp = document.getElementById("snsLblIp");
        var lblPort = document.getElementById("snsLblPort");
        var remoteOn = !local;
        if (comEl) comEl.disabled = !local;
        if (ipEl) ipEl.disabled = !remoteOn;
        if (portEl) portEl.disabled = !remoteOn;
        if (portLink) {
            portLink.style.display = local ? "none" : "inline";
            portLink.disabled = !remoteOn;
            portLink.tabIndex = remoteOn ? 0 : -1;
        }
        if (lblIp) lblIp.classList.toggle("sns-ctl-disabled", !remoteOn);
        if (lblPort) lblPort.classList.toggle("sns-ctl-disabled", !remoteOn);

        var commGrp = document.getElementById("snsCommGroup");
        if (commGrp) commGrp.disabled = !local;

        var baud = document.getElementById("snsBaud");
        var db = document.getElementById("snsDataBits");
        var par = document.getElementById("snsParity");
        var stb = document.getElementById("snsStopBits");

        var advGrp = document.getElementById("snsAdvGroup");
        var rowStd = document.getElementById("snsRowStatusDelay");
        var rowDrop = document.getElementById("snsRowDrop");
        var rowTx = document.getElementById("snsRowTransmit");
        var rowStuck = document.getElementById("snsRowStuck");
        var advWarn = document.getElementById("snsAdvWarn");

        var spl = p === "SPL-100 Push" || p === "SPL-200 Push";
        if (advGrp) advGrp.style.display = spl ? "none" : "";
        if (advWarn) advWarn.style.display = spl ? "none" : "";
        var snShell = document.getElementById("snSetupShell");
        if (snShell) snShell.classList.toggle("sns-compact", spl);

        if (rowStd) rowStd.style.display = p === "Protocol A" ? "" : "none";
        if (rowDrop) rowDrop.style.display = p === "Protocol B" ? "" : "none";
        if (rowTx) rowTx.style.display = p === "Modbus/RTU" || p === "HART Protocol" ? "" : "none";
        if (rowStuck) rowStuck.style.display = p === "Protocol B" ? "" : "none";

        if (!local) return;

        if (spl) {
            if (baud) baud.disabled = false;
            if (db) {
                db.disabled = true;
                db.value = "8";
            }
            if (par) {
                par.disabled = true;
                par.value = "None";
            }
            if (stb) {
                stb.disabled = true;
                stb.value = "1";
            }
            return;
        }

        function lockComm(a, b, c, d, enB, enDb, enP, enS) {
            if (a != null && baud) baud.value = String(a);
            if (b != null && db) db.value = String(b);
            if (c != null && par) {
                par.value = c === "N" || c === "None" ? "None" : c === "E" || c === "Even" ? "Even" : "Odd";
            }
            if (d != null && stb) stb.value = String(d);
            if (baud) baud.disabled = !enB;
            if (db) db.disabled = !enDb;
            if (par) par.disabled = !enP;
            if (stb) stb.disabled = !enS;
        }

        switch (p) {
            case "Protocol A":
                lockComm(2400, 8, "None", 1, false, false, false, false);
                break;
            case "Protocol B":
                lockComm(2400, 8, "None", 1, false, false, false, false);
                break;
            case "Modbus/RTU":
                lockComm(null, 8, null, null, true, false, true, true);
                if (db) db.value = "8";
                break;
            case "HART Protocol":
                lockComm(null, 8, null, 1, true, false, true, false);
                if (db) db.value = "8";
                break;
            default:
                lockComm(null, null, null, null, true, false, true, true);
                break;
        }
    }

    function openSensorNetworkSetup(networkId) {
        var n = findNetwork(networkId);
        if (!n) return;
        ensureNetworkFields(n);
        snSetupEditingId = networkId;
        var local = n.interfaceMode !== "remote";
        var comOpts = "";
        var comSeen = {};
        var c;
        for (c = 1; c <= 6; c++) {
            var com = "COM" + c;
            comSeen[com] = true;
            comOpts +=
                "<option value=\"" +
                com +
                "\"" +
                (n.interface === com ? " selected" : "") +
                ">" +
                com +
                "</option>";
        }
        if (n.interface && !comSeen[n.interface]) {
            comOpts +=
                "<option value=\"" +
                escapeHtml(n.interface) +
                "\" selected>" +
                escapeHtml(n.interface) +
                "</option>";
        }
        var html =
            '<div class="sns-frm">' +
            '<div class="sns-heading">Sensor Network Setup</div>' +
            '<fieldset class="sns-group" id="snsGrpName"><legend>Name / Description</legend>' +
            '<input type="text" class="sns-input-wide" id="snsName" value="' +
            escapeHtml(n.name) +
            '">' +
            "</fieldset>" +
            '<fieldset class="sns-group" id="snsGrpProto"><legend>Protocol</legend>' +
            '<label class="sns-proto-lbl" for="snsProtocol">Protocol used by all sensors on this network:</label>' +
            '<select id="snsProtocol" class="sns-input-protocol">' +
            buildProtocolOptionsHtml(n.protocol) +
            "</select>" +
            "</fieldset>" +
            '<fieldset class="sns-group" id="snsGrpIface"><legend>Interface</legend>' +
            '<div class="sns-radio-block">' +
            "<label>" +
            '<input type="radio" name="snsIf" id="snsIfLocal" value="local"' +
            (local ? " checked" : "") +
            "> Local Serial Port:</label>" +
            '<select id="snsCom">' +
            comOpts +
            "</select>" +
            "</div>" +
            '<div class="sns-radio-block">' +
            "<label>" +
            '<input type="radio" name="snsIf" id="snsIfRemote" value="remote"' +
            (!local ? " checked" : "") +
            "> Remote Serial Server:</label>" +
            '<div class="sns-remote-stack">' +
            '<div class="sns-remote-line">' +
            '<label class="sns-if-sublbl" id="snsLblIp" for="snsIp">IP Address:</label>' +
            '<input type="text" id="snsIp" class="sns-if-input sns-ip-inp" value="' +
            escapeHtml(n.remoteIp) +
            '">' +
            "</div>" +
            '<div class="sns-remote-line sns-remote-line-port">' +
            '<label class="sns-if-sublbl" id="snsLblPort" for="snsPort">Port Number:</label>' +
            '<input type="text" id="snsPort" class="sns-if-input sns-port-inp" value="' +
            escapeHtml(n.remotePort) +
            '">' +
            '<button type="button" class="sns-link-btn" id="snsPortLink" title="Common serial server ports">Common Port Numbers</button>' +
            "</div>" +
            "</div>" +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="sns-group" id="snsCommGroup"><legend>Communication Parameters</legend>' +
            '<div class="sns-comm-grid">' +
            "<div><span>Baud Rate</span><select id=\"snsBaud\" class=\"sns-comm-select\">" +
            buildBaudOptionsHtml(n.baud) +
            "</select></div>" +
            "<div><span>Data Bits</span><select id=\"snsDataBits\" class=\"sns-comm-select\">" +
            buildDataBitsOptionsHtml(n.dataBits) +
            "</select></div>" +
            "<div><span>Parity</span><select id=\"snsParity\" class=\"sns-comm-select\">" +
            buildParityOptionsHtml(n.parity) +
            "</select></div>" +
            "<div><span>Stop Bits</span><select id=\"snsStopBits\" class=\"sns-comm-select\">" +
            buildStopBitsOptionsHtml(n.stopBits) +
            "</select></div>" +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="sns-group" id="snsAdvGroup"><legend>Advanced</legend>' +
            '<p id="snsAdvWarn" class="sns-adv-hint">These advanced controls should only be adjusted under the guidance of BinMaster technical support.</p>' +
            '<div class="sns-adv-base">' +
            "<label><span>Request Attempts</span><input type=\"number\" id=\"snsAtt\" step=\"1\" min=\"1\" max=\"10\" value=\"" +
            n.requestAttempts +
            '"></label>' +
            "<label><span>Response Timeout</span><input type=\"number\" id=\"snsRto\" step=\"0.1\" min=\"0.1\" max=\"20\" value=\"" +
            n.responseTimeout +
            '"></label>' +
            "</div>" +
            '<div class="sns-adv-extra" id="snsRowStatusDelay">' +
            "<label><span>Status Request Delay</span><input type=\"number\" id=\"snsStatusDelay\" step=\"0.1\" min=\"0\" max=\"5\" value=\"" +
            n.statusRequestDelay +
            '"></label>' +
            "</div>" +
            '<div class="sns-adv-extra" id="snsRowDrop">' +
            "<label><span>Drop Timeout</span><input type=\"number\" id=\"snsDropTimeout\" step=\"1\" min=\"20\" max=\"200\" value=\"" +
            n.dropTimeout +
            '"></label>' +
            "</div>" +
            '<div class="sns-adv-extra" id="snsRowTransmit">' +
            "<label><span>Transmit Delay</span><input type=\"number\" id=\"snsTransmitDelay\" step=\"1\" min=\"5\" max=\"500\" value=\"" +
            n.transmitDelay +
            '"></label>' +
            "</div>" +
            '<div class="sns-adv-extra sns-adv-stuck" id="snsRowStuck">' +
            "<label class=\"sns-stuck-lbl\"><span>Disable StuckTop</span>" +
            '<input type="checkbox" id="snsDisableStuckTop"' +
            (n.disableStuckTop ? " checked" : "") +
            "></label>" +
            "</div>" +
            "</fieldset>" +
            "</div>";

        document.getElementById("snSetupBody").innerHTML = html;
        document.getElementById("backdropSnSetup").classList.add("show");

        document.getElementById("snsIfLocal").addEventListener("change", function () {
            syncSensorNetworkSetupUi(false);
        });
        document.getElementById("snsIfRemote").addEventListener("change", function () {
            syncSensorNetworkSetupUi(false);
        });
        document.getElementById("snsProtocol").addEventListener("change", function () {
            syncSensorNetworkSetupUi(true);
        });
        document.getElementById("snsPortLink").addEventListener("click", function () {
            var portEl = document.getElementById("snsPort");
            if (portEl && !portEl.disabled) {
                portEl.value = "50000";
            }
        });
        syncSensorNetworkSetupUi(false);
    }

    function saveSensorNetworkSetup() {
        var n = findNetwork(snSetupEditingId);
        if (!n) {
            closeSensorNetworkSetup();
            return;
        }
        n.name = document.getElementById("snsName").value.trim() || n.name;
        n.protocol = document.getElementById("snsProtocol").value;
        n.interfaceMode = document.getElementById("snsIfRemote").checked ? "remote" : "local";
        if (n.interfaceMode === "local") {
            n.interface = document.getElementById("snsCom").value;
        } else {
            n.remoteIp = document.getElementById("snsIp").value.trim();
            n.remotePort = document.getElementById("snsPort").value.trim();
        }
        n.baud = parseInt(document.getElementById("snsBaud").value, 10) || 2400;
        n.dataBits = parseInt(document.getElementById("snsDataBits").value, 10) || 8;
        n.parity = document.getElementById("snsParity").value;
        n.stopBits = parseInt(document.getElementById("snsStopBits").value, 10) || 1;
        n.requestAttempts = parseInt(document.getElementById("snsAtt").value, 10) || 3;
        n.responseTimeout = parseFloat(document.getElementById("snsRto").value) || 1.5;
        var stdEl = document.getElementById("snsStatusDelay");
        var dropEl = document.getElementById("snsDropTimeout");
        var txEl = document.getElementById("snsTransmitDelay");
        var stkEl = document.getElementById("snsDisableStuckTop");
        if (stdEl) n.statusRequestDelay = parseFloat(stdEl.value) || 1;
        if (dropEl) n.dropTimeout = parseInt(dropEl.value, 10) || 120;
        n.transmitDelay = txEl ? parseFloat(txEl.value) : n.transmitDelay;
        if (isNaN(n.transmitDelay)) n.transmitDelay = 20;
        if (stkEl) n.disableStuckTop = stkEl.checked;
        syncCommParamsString(n);
        saveState();
        closeSensorNetworkSetup();
        toast("Sensor network saved (simulation).");
        if (document.getElementById("backdropApp").classList.contains("show")) {
            var body = document.getElementById("appModalBody");
            if (body && body.querySelector(".sn-win")) {
                openSensorNetworks();
            }
        }
    }

    function openSensorNetworks() {
        state.sensorNetworks.forEach(ensureNetworkFields);
        saveState();

        var validSel =
            snSelectedNetworkId && state.sensorNetworks.some(function (n) { return n.id === snSelectedNetworkId; });
        if (!validSel) snSelectedNetworkId = state.sensorNetworks[0] ? state.sensorNetworks[0].id : null;
        var canAct = !!snSelectedNetworkId && state.sensorNetworks.length > 0;

        var rows = state.sensorNetworks
            .map(function (n) {
                var sel = n.id === snSelectedNetworkId ? " class=\"sn-selected\"" : "";
                return (
                    "<tr data-nid=\"" +
                    escapeHtml(n.id) +
                    "\"" +
                    sel +
                    ">" +
                    "<td>" +
                    escapeHtml(n.name) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(n.protocol) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(networkInterfaceDisplayForTable(n)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(n.commParams) +
                    "</td>" +
                    "</tr>"
                );
            })
            .join("");

        var html =
            '<div class="sn-win">' +
            '<div class="sn-title">Sensor Networks</div>' +
            '<div class="sn-layout">' +
            '<div class="sn-table-wrap">' +
            '<table class="sn-table">' +
            "<thead><tr><th>Network Name</th><th>Protocol</th><th>Interface</th><th>Communication Parameters</th></tr></thead>" +
            "<tbody>" +
            (rows || '<tr><td colspan="4">&nbsp;</td></tr>') +
            "</tbody>" +
            "</table>" +
            "</div>" +
            '<div class="sn-actions">' +
            '<button type="button" id="snBtnSelect"' +
            (canAct ? "" : " disabled") +
            ">Select</button>" +
            '<button type="button" id="snAddNew">Add New</button>' +
            '<button type="button" id="snBtnDelete"' +
            (canAct ? "" : " disabled") +
            ">Delete</button>" +
            '<button type="button" id="snBtnDevice"' +
            (canAct ? "" : " disabled") +
            ">Device Interface</button>" +
            '<button type="button" class="sn-close" data-close-app>Close</button>' +
            "</div>" +
            "</div>" +
            "</div>";

        openAppModal("Sensor Networks - Binventory Workstation", html, "", "modal-sn-networks");

        function setSnActionButtons(enabled) {
            ["snBtnSelect", "snBtnDelete", "snBtnDevice"].forEach(function (id) {
                var b = document.getElementById(id);
                if (b) b.disabled = !enabled;
            });
        }

        appModalBody.querySelectorAll(".sn-table tbody tr[data-nid]").forEach(function (tr) {
            tr.addEventListener("click", function (e) {
                e.stopPropagation();
                snSelectedNetworkId = tr.getAttribute("data-nid");
                appModalBody.querySelectorAll(".sn-table tbody tr[data-nid]").forEach(function (r) {
                    r.classList.remove("sn-selected");
                });
                tr.classList.add("sn-selected");
                setSnActionButtons(true);
            });
            tr.addEventListener("dblclick", function (e) {
                e.preventDefault();
                e.stopPropagation();
                snSelectedNetworkId = tr.getAttribute("data-nid");
                openSensorNetworkSetup(snSelectedNetworkId);
            });
        });

        var btnSel = document.getElementById("snBtnSelect");
        if (btnSel) {
            btnSel.addEventListener("click", function () {
                if (snSelectedNetworkId) openSensorNetworkSetup(snSelectedNetworkId);
            });
        }
        var btnDel = document.getElementById("snBtnDelete");
        if (btnDel) {
            btnDel.addEventListener("click", function () {
                if (!snSelectedNetworkId || state.sensorNetworks.length <= 1) {
                    toast("Keep at least one sensor network.");
                    return;
                }
                state.sensorNetworks = state.sensorNetworks.filter(function (x) {
                    return x.id !== snSelectedNetworkId;
                });
                snSelectedNetworkId = state.sensorNetworks[0] ? state.sensorNetworks[0].id : null;
                saveState();
                closeAppModal();
                openSensorNetworks();
                toast("Network deleted (simulation).");
            });
        }
        var btnDev = document.getElementById("snBtnDevice");
        if (btnDev) {
            btnDev.addEventListener("click", function () {
                if (snSelectedNetworkId) openSensorNetworkSetup(snSelectedNetworkId);
            });
        }

        document.getElementById("snAddNew").addEventListener("click", function () {
            var k = state.sensorNetworks.length + 1;
            var comNum = Math.min(6, Math.max(1, k));
            var newId = uid("n");
            state.sensorNetworks.push({
                id: newId,
                name: "COM" + comNum + " Sensor Network",
                protocol: "Protocol A",
                interface: "COM" + comNum,
                commParams: "2400,8,N,1",
                siteId: state.sites && state.sites[0] ? state.sites[0].id : undefined
            });
            snSelectedNetworkId = newId;
            saveState();
            closeAppModal();
            openSensorNetworks();
            toast("Network added (simulation).");
        });
    }

    /** frmEmailSetup.vb / frmEmailSetup.Designer.vb — ClientSize 509×573, control order top to bottom. */
    function ensureEmailSettingsDefaults() {
        if (!state.emailSettings) state.emailSettings = {};
        var e = state.emailSettings;
        if (e.enableSmtpClient == null) e.enableSmtpClient = true;
        if (e.smtp == null) e.smtp = "";
        if (e.port == null) e.port = "25";
        if (e.useSecureConnection == null) e.useSecureConnection = false;
        if (e.useSmtpAuth == null) e.useSmtpAuth = false;
        if (e.smtpUserId == null) e.smtpUserId = "";
        if (e.smtpPassword == null) e.smtpPassword = "";
        if (e.smtpVerifyPassword == null) e.smtpVerifyPassword = "";
        if (e.adminEmail == null) e.adminEmail = "admin@ebob.com";
        if (e.defaultFrom == null) {
            e.defaultFrom =
                e.fromAddr != null && e.fromAddr !== "" ? e.fromAddr : "messenger@ebob.com";
        }
        if (e.fromAddr == null) e.fromAddr = e.defaultFrom;
        if (e.adminSubject == null) e.adminSubject = "Administrative Alert";
        if (e.measureSubject == null) e.measureSubject = "Vessel Status";
        if (e.alarmSubject == null) e.alarmSubject = "Vessel Alarm";
        if (e.emailAppend == null) e.emailAppend = "";
    }

    function openEmailSetup() {
        ensureEmailSettingsDefaults();
        var e = state.emailSettings;
        var enChk = e.enableSmtpClient ? " checked" : "";
        var secChk = e.useSecureConnection ? " checked" : "";
        var authChk = e.useSmtpAuth ? " checked" : "";
        var html =
            '<div class="es-frm">' +
            '<div class="es-lbl-title">Email Setup</div>' +
            '<label class="es-chk-enable"><input type="checkbox" id="esEnable"' +
            enChk +
            '> Enable SMTP Email Client</label>' +
            '<fieldset class="win-groupbox es-gb js-es-dep" id="esGbSmtp">' +
            "<legend>SMTP Server</legend>" +
            '<div class="es-field-row">' +
            '<label for="esServer">Server Name:</label>' +
            '<input type="text" id="esServer" class="es-input-long" value="' +
            escapeHtml(e.smtp) +
            '" autocomplete="off">' +
            "</div>" +
            '<div class="es-field-row es-smtp-port-row">' +
            '<label for="esPort">SMTP Port:</label>' +
            '<input type="text" id="esPort" class="es-input-port" value="' +
            escapeHtml(e.port) +
            '" inputmode="numeric" autocomplete="off">' +
            '<span class="es-port-flex"></span>' +
            '<label class="es-chk-secure"><input type="checkbox" id="esSecure"' +
            secChk +
            '> Use a Secure Connection</label>' +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="win-groupbox es-gb es-gb-auth js-es-dep" id="esGbAuth">' +
            "<legend>SMTP Authentication</legend>" +
            '<label class="es-chk-auth"><input type="checkbox" id="esAuth"' +
            authChk +
            '> Use SMTP Authentication</label>' +
            '<div class="es-field-row es-auth-row">' +
            '<label for="esUser">User ID:</label>' +
            '<input type="text" id="esUser" class="es-input-long" value="' +
            escapeHtml(e.smtpUserId) +
            '" autocomplete="off">' +
            "</div>" +
            '<div class="es-field-row es-auth-row">' +
            '<label for="esPass">Password:</label>' +
            '<input type="password" id="esPass" class="es-input-long" value="' +
            escapeHtml(e.smtpPassword) +
            '" autocomplete="new-password">' +
            "</div>" +
            '<div class="es-field-row es-auth-row">' +
            '<label for="esVerify">Verify Password:</label>' +
            '<input type="password" id="esVerify" class="es-input-long" value="' +
            escapeHtml(e.smtpVerifyPassword) +
            '" autocomplete="new-password">' +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="win-groupbox es-gb js-es-dep" id="esGbAddr">' +
            "<legend>Email Addresses</legend>" +
            '<div class="es-field-row">' +
            '<label for="esAdmin">Administrator:</label>' +
            '<input type="text" id="esAdmin" class="es-input-long" value="' +
            escapeHtml(e.adminEmail) +
            '" autocomplete="off">' +
            "</div>" +
            '<div class="es-field-row">' +
            '<label for="esFrom">Default From:</label>' +
            '<input type="text" id="esFrom" class="es-input-long" value="' +
            escapeHtml(e.defaultFrom) +
            '" autocomplete="off">' +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="win-groupbox es-gb js-es-dep" id="esGbSubj">' +
            "<legend>Message Subject Line</legend>" +
            '<div class="es-field-row">' +
            '<label for="esSubjAdmin">Administrative:</label>' +
            '<input type="text" id="esSubjAdmin" class="es-input-long" value="' +
            escapeHtml(e.adminSubject) +
            '">' +
            "</div>" +
            '<div class="es-field-row">' +
            '<label for="esSubjMeas">Measurement:</label>' +
            '<input type="text" id="esSubjMeas" class="es-input-long" value="' +
            escapeHtml(e.measureSubject) +
            '">' +
            "</div>" +
            '<div class="es-field-row">' +
            '<label for="esSubjAlarm">Alarm:</label>' +
            '<input type="text" id="esSubjAlarm" class="es-input-long" value="' +
            escapeHtml(e.alarmSubject) +
            '">' +
            "</div>" +
            "</fieldset>" +
            '<fieldset class="win-groupbox es-gb es-gb-trailer js-es-dep" id="esGbTrailer">' +
            "<legend>Email Trailer</legend>" +
            '<p class="es-trailer-hint">Define a standard trailer here to append to all email messages generated by the system.</p>' +
            '<textarea id="esAppend" class="es-textarea-append" rows="3" spellcheck="false">' +
            escapeHtml(e.emailAppend) +
            "</textarea>" +
            "</fieldset>" +
            "</div>";
        openAppModal(
            "Email Setup - Binventory Workstation",
            html,
            '<button type="button" id="esSave" class="win-btn win-btn-default" accesskey="s">Save</button>' +
                '<button type="button" class="win-btn" data-close-app accesskey="c">Cancel</button>',
            "modal-email-setup modal-win-toolwindow"
        );

        function syncEsAuthFields() {
            var master = document.getElementById("esEnable");
            var authEl = document.getElementById("esAuth");
            if (!master || !master.checked) return;
            var useAuth = authEl && authEl.checked;
            ["esUser", "esPass", "esVerify"].forEach(function (id) {
                var inp = document.getElementById(id);
                if (inp) inp.disabled = !useAuth;
            });
        }

        function syncEsGroups() {
            var master = document.getElementById("esEnable");
            var on = master && master.checked;
            document.querySelectorAll(".es-frm .js-es-dep").forEach(function (fs) {
                if (fs && fs.tagName === "FIELDSET") fs.disabled = !on;
            });
            if (on) syncEsAuthFields();
        }

        var esEnable = document.getElementById("esEnable");
        var esAuth = document.getElementById("esAuth");
        if (esEnable) {
            esEnable.addEventListener("change", syncEsGroups);
        }
        if (esAuth) {
            esAuth.addEventListener("change", syncEsAuthFields);
        }
        syncEsGroups();

        var esSave = document.getElementById("esSave");
        if (esSave) {
            esSave.addEventListener("click", function () {
                var pass = document.getElementById("esPass").value;
                var ver = document.getElementById("esVerify").value;
                if (document.getElementById("esAuth").checked && pass !== ver) {
                    toast("Password and Verify Password do not match.");
                    return;
                }
                e.enableSmtpClient = document.getElementById("esEnable").checked;
                e.smtp = document.getElementById("esServer").value.trim();
                e.port = document.getElementById("esPort").value.trim();
                e.useSecureConnection = document.getElementById("esSecure").checked;
                e.useSmtpAuth = document.getElementById("esAuth").checked;
                e.smtpUserId = document.getElementById("esUser").value.trim();
                e.smtpPassword = pass;
                e.smtpVerifyPassword = ver;
                e.adminEmail = document.getElementById("esAdmin").value.trim();
                e.defaultFrom = document.getElementById("esFrom").value.trim();
                e.fromAddr = e.defaultFrom;
                e.adminSubject = document.getElementById("esSubjAdmin").value.trim();
                e.measureSubject = document.getElementById("esSubjMeas").value.trim();
                e.alarmSubject = document.getElementById("esSubjAlarm").value.trim();
                e.emailAppend = document.getElementById("esAppend").value;
                saveState();
                closeAppModal();
                toast("Email settings saved (simulation).");
            });
        }

        setTimeout(function () {
            var el = document.getElementById("esServer");
            if (el && !el.disabled) el.focus();
        }, 0);
    }

    function ensureSystemSettingsDefaults() {
        if (!state.systemSettings) state.systemSettings = {};
        var s = state.systemSettings;
        if (s.units == null) s.units = "Imperial";
        if (s.timezone == null) s.timezone = "US/Central";
        if (s.registeredUser == null) s.registeredUser = "";
        if (s.companyName == null) s.companyName = "";
        if (s.streetAddress == null) s.streetAddress = "";
        if (s.streetAddress2 == null) s.streetAddress2 = "";
        if (s.city == null) s.city = "";
        if (s.state == null) s.state = "";
        if (s.zipCode == null) s.zipCode = "";
        if (s.country == null) s.country = "";
        if (s.ldapAddress == null) s.ldapAddress = "";
        if (s.measurementRetentionDays == null || String(s.measurementRetentionDays).trim() === "") {
            s.measurementRetentionDays = "60";
        }
        if (s.autoLogin == null) s.autoLogin = true;
    }

    /** frmSystem — Registered User & Company + System Options (ClientSize 669×424). */
    function openSystemSetup() {
        ensureSystemSettingsDefaults();
        state.sites.forEach(ensureSiteFields);
        var s = state.systemSettings;
        var autoChk = s.autoLogin ? " checked" : "";
        var html =
            '<div class="ss-form">' +
            '<div class="ss-form-title">System Setup</div>' +
            '<fieldset class="ss-fieldset">' +
            "<legend>Registered User &amp; Company</legend>" +
            '<div class="ss-row"><label for="ssRegUser">Registered User:</label>' +
            '<input type="text" id="ssRegUser" class="ss-input-full" value="' +
            escapeHtml(s.registeredUser) +
            '"></div>' +
            '<div class="ss-row"><label for="ssCompany">Company Name:</label>' +
            '<input type="text" id="ssCompany" class="ss-input-full" value="' +
            escapeHtml(s.companyName) +
            '"></div>' +
            '<div class="ss-row"><label for="ssStreet1">Street Address:</label>' +
            '<input type="text" id="ssStreet1" class="ss-input-full" value="' +
            escapeHtml(s.streetAddress) +
            '"></div>' +
            '<div class="ss-row ss-row-addr2"><label class="ss-lbl-empty" for="ssStreet2">&nbsp;</label>' +
            '<input type="text" id="ssStreet2" class="ss-input-full" value="' +
            escapeHtml(s.streetAddress2) +
            '"></div>' +
            '<div class="ss-row"><label for="ssCity">City/Town:</label>' +
            '<input type="text" id="ssCity" class="ss-input-full" value="' +
            escapeHtml(s.city) +
            '"></div>' +
            '<div class="ss-row"><label for="ssState">State/Province/Region:</label>' +
            '<input type="text" id="ssState" class="ss-input-full" value="' +
            escapeHtml(s.state) +
            '"></div>' +
            '<div class="ss-row"><label for="ssZip">Zip/Postal Code:</label>' +
            '<input type="text" id="ssZip" class="ss-input-zip" value="' +
            escapeHtml(s.zipCode) +
            '"></div>' +
            '<div class="ss-row"><label for="ssCountry">Country/Nation:</label>' +
            '<input type="text" id="ssCountry" class="ss-input-full" value="' +
            escapeHtml(s.country) +
            '"></div>' +
            "</fieldset>" +
            '<fieldset class="ss-fieldset ss-fieldset-system">' +
            "<legend>System Options</legend>" +
            '<div class="ss-row"><label for="ssLdap">LDAP Server Address:</label>' +
            '<input type="text" id="ssLdap" class="ss-input-full" value="' +
            escapeHtml(s.ldapAddress) +
            '"></div>' +
            '<div class="ss-row"><label for="ssRetention">Measurement Retention Days:</label>' +
            '<input type="text" id="ssRetention" class="ss-input-retention" value="' +
            escapeHtml(String(s.measurementRetentionDays)) +
            '"></div>' +
            '<div class="ss-row ss-row-check">' +
            '<span class="ss-check-text">Do NOT require users to log in:</span>' +
            '<input type="checkbox" id="ssAutoLogin"' +
            autoChk +
            "></div>" +
            "</fieldset>" +
            "</div>";
        openAppModal(
            "System Setup - Binventory Workstation",
            html,
            '<button type="button" id="ssSave" class="primary ss-footer-btn">Save</button><button type="button" class="secondary ss-footer-btn" data-close-app>Cancel</button>',
            "modal-system-setup"
        );
        document.getElementById("ssSave").addEventListener("click", function () {
            s.registeredUser = document.getElementById("ssRegUser").value;
            s.companyName = document.getElementById("ssCompany").value;
            s.streetAddress = document.getElementById("ssStreet1").value;
            s.streetAddress2 = document.getElementById("ssStreet2").value;
            s.city = document.getElementById("ssCity").value;
            s.state = document.getElementById("ssState").value;
            s.zipCode = document.getElementById("ssZip").value;
            s.country = document.getElementById("ssCountry").value;
            s.ldapAddress = document.getElementById("ssLdap").value;
            s.measurementRetentionDays = document.getElementById("ssRetention").value.trim() || "60";
            s.autoLogin = document.getElementById("ssAutoLogin").checked;
            saveState();
            closeAppModal();
            toast("System settings saved.");
        });
    }

    function openChangePassword() {
        openAppModal(
            "Change Password",
            '<div class="form-grid">' +
            "<label>New password<br><input type=\"password\" id=\"p1\"></label>" +
            "<label>Confirm<br><input type=\"password\" id=\"p2\"></label>" +
            "</div>",
            '<button type="button" id="pwOk" class="primary">OK</button><button type="button" class="secondary" data-close-app>Cancel</button>'
        );
        document.getElementById("pwOk").addEventListener("click", function () {
            if (document.getElementById("p1").value !== document.getElementById("p2").value) {
                toast("Passwords do not match.");
                return;
            }
            closeAppModal();
            toast("Password updated (simulation only).");
        });
    }

    /**
     * frmUserMaintenance.vb — DataGridView columns: User ID, User Type, FirstName, Middle Init, Last Name, Last Logon;
     * buttons top→bottom: Select, Add New, Delete; Close at bottom (Designer ClientSize 859×336).
     */
    function ensureUserMaintenanceUser(u) {
        if (!u || typeof u !== "object") return u;
        if (u.userId == null || String(u.userId).trim() === "") u.userId = u.id || u.name || "User";
        if (u.userType == null) u.userType = u.role || "";
        if (u.firstName == null) u.firstName = "";
        if (u.middleInit == null) u.middleInit = "";
        if (u.lastName == null) u.lastName = "";
        if (u.lastLogon == null) u.lastLogon = "";
        if (u.firstName === "" && u.lastName === "" && u.name) {
            var parts = String(u.name).trim().split(/\s+/);
            if (parts.length === 1) {
                u.firstName = parts[0];
            } else if (parts.length > 1) {
                u.firstName = parts[0];
                u.lastName = parts.slice(1).join(" ");
            }
        }
        return u;
    }

    function loggedInUserIdString() {
        var rec = state.users.filter(function (x) {
            return x.name === state.currentUser;
        })[0];
        if (rec && rec.userId != null) return String(rec.userId);
        return state.currentUser ? String(state.currentUser) : "";
    }

    function openUserMaintenance() {
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to open User Maintenance.");
            return;
        }
        state.users.forEach(ensureUserMaintenanceUser);
        var rowsHtml = state.users
            .map(function (u, idx) {
                return (
                    '<tr data-um-idx="' +
                    idx +
                    '" tabindex="-1">' +
                    "<td>" +
                    escapeHtml(String(u.userId)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(String(u.userType)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(String(u.firstName)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(String(u.middleInit)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(String(u.lastName)) +
                    "</td>" +
                    "<td>" +
                    escapeHtml(String(u.lastLogon)) +
                    "</td>" +
                    "</tr>"
                );
            })
            .join("");
        var html =
            '<div class="um-win">' +
            '<div class="um-lbl-title">User Maintenance</div>' +
            '<div class="um-main-row">' +
            '<div class="um-dgv-wrap" role="presentation">' +
            '<table class="win-dgv um-dgv" tabindex="0">' +
            "<thead><tr>" +
            "<th>User ID</th>" +
            "<th>User Type</th>" +
            "<th>FirstName</th>" +
            "<th>Middle Init</th>" +
            "<th>Last Name</th>" +
            "<th>Last Logon</th>" +
            "</tr></thead>" +
            "<tbody>" +
            rowsHtml +
            "</tbody></table></div>" +
            '<div class="um-actions-col">' +
            '<button type="button" class="win-btn win-btn-default" id="umSelect" disabled accesskey="s">Select</button>' +
            '<button type="button" class="win-btn" id="umAddNew" accesskey="a">Add New</button>' +
            '<button type="button" class="win-btn" id="umDelete" disabled accesskey="d">Delete</button>' +
            '<span class="um-actions-spacer" aria-hidden="true"></span>' +
            '<button type="button" class="win-btn" data-close-app accesskey="c">Close</button>' +
            "</div></div></div>";

        openAppModal(
            "User Maintenance - Binventory Workstation",
            html,
            "",
            "modal-user-maintenance modal-footer-hidden modal-win-toolwindow"
        );

        var tbody = document.querySelector(".um-dgv tbody");
        var selBtn = document.getElementById("umSelect");
        var delBtn = document.getElementById("umDelete");
        var selectedIdx = -1;

        function clearSelection() {
            selectedIdx = -1;
            if (tbody) {
                tbody.querySelectorAll("tr").forEach(function (tr) {
                    tr.classList.remove("selected");
                });
            }
            if (selBtn) selBtn.disabled = true;
            if (delBtn) delBtn.disabled = true;
        }

        function applySelection(idx) {
            if (!tbody || idx < 0 || idx >= state.users.length) {
                clearSelection();
                return;
            }
            selectedIdx = idx;
            tbody.querySelectorAll("tr").forEach(function (tr, i) {
                tr.classList.toggle("selected", i === idx);
            });
            if (selBtn) selBtn.disabled = false;
            if (delBtn) delBtn.disabled = false;
        }

        function doOpenSelection() {
            if (selectedIdx < 0 || selectedIdx >= state.users.length) return;
            showBinventoryMessageBox({
                icon: "info",
                message: "User setup (simulation) — edit user is not implemented in this tutorial.",
                buttons: "ok"
            });
        }

        if (tbody) {
            tbody.addEventListener("click", function (e) {
                var tr = e.target.closest("tr[data-um-idx]");
                if (!tr) return;
                e.preventDefault();
                applySelection(parseInt(tr.getAttribute("data-um-idx"), 10));
            });
            tbody.addEventListener("dblclick", function (e) {
                var tr = e.target.closest("tr[data-um-idx]");
                if (!tr) return;
                e.preventDefault();
                applySelection(parseInt(tr.getAttribute("data-um-idx"), 10));
                doOpenSelection();
            });
        }

        var sel = document.getElementById("umSelect");
        if (sel) sel.addEventListener("click", doOpenSelection);

        var addNew = document.getElementById("umAddNew");
        if (addNew) {
            addNew.addEventListener("click", function () {
                var n = state.users.length + 1;
                var uidStr = "User" + n;
                var guard = 0;
                while (state.users.some(function (u) { return String(u.userId) === uidStr; }) && guard < 500) {
                    n += 1;
                    uidStr = "User" + n;
                    guard += 1;
                }
                state.users.push({
                    id: uid("u"),
                    userId: uidStr,
                    name: "User " + n,
                    role: "Operator",
                    userType: "Operator",
                    firstName: "User",
                    middleInit: "",
                    lastName: String(n),
                    lastLogon: ""
                });
                saveState();
                closeAppModal();
                openUserMaintenance();
            });
        }

        if (delBtn) {
            delBtn.addEventListener("click", function () {
                if (selectedIdx < 0 || selectedIdx >= state.users.length) return;
                var u = state.users[selectedIdx];
                var uid = String(u.userId);
                if (uid === loggedInUserIdString()) {
                    showBinventoryMessageBox({
                        icon: "warn",
                        message:
                            "You are currently logged in as this user. You must be logged in as a different administrator to delete this user.",
                        buttons: "ok"
                    });
                    return;
                }
                showBinventoryMessageBox({
                    icon: "warn",
                    message: "Are you sure you want to delete the user '" + uid + "'?",
                    buttons: "okcancel",
                    onOk: function () {
                        state.users.splice(selectedIdx, 1);
                        saveState();
                        closeAppModal();
                        openUserMaintenance();
                    }
                });
            });
        }

        clearSelection();
    }

    /**
     * frmEmailReports.vb — frmEmailReports_Load sets Text; form has NO Save (only btnOk &Select, btnCancel &Close).
     * lstReports.Items: "Site Status Report" only. AcceptButton=btnOk, CancelButton=btnCancel.
     * btnOk_Click: empty selection → MsgBox "Please select a report."; Case "Site Status Report" → frmScheduleMaintenance miEventType=3;
     * Case Else → MsgBox "Invalid Report." Exclamation.
     */
    function openEmailReports() {
        var html =
            '<div class="er-binv">' +
            '<div class="er-binv-title">Email Reports</div>' +
            '<div class="er-binv-main">' +
            '<div class="er-listbox" id="erLstReports" role="listbox" tabindex="0" aria-label="Reports" aria-multiselectable="false">' +
            '<div class="er-opt" role="option" id="erOptSite" data-report="Site Status Report" aria-selected="false">Site Status Report</div>' +
            "</div>" +
            '<div class="er-binv-actions">' +
            '<button type="button" class="win-btn win-btn-default" id="erBtnSelect" accesskey="s">Select</button>' +
            '<button type="button" class="win-btn" id="erBtnClose" data-close-app accesskey="c">Close</button>' +
            "</div></div></div>";

        openAppModal(
            "Email Reports - Binventory Workstation",
            html,
            "",
            "modal-email-reports modal-footer-hidden modal-win-toolwindow"
        );

        var lb = document.getElementById("erLstReports");
        var optEl = document.getElementById("erOptSite");
        var selectedReport = null;

        function clearSelection() {
            selectedReport = null;
            if (optEl) {
                optEl.classList.remove("er-opt-selected");
                optEl.setAttribute("aria-selected", "false");
            }
        }

        function selectSiteStatus() {
            selectedReport = "Site Status Report";
            if (optEl) {
                optEl.classList.add("er-opt-selected");
                optEl.setAttribute("aria-selected", "true");
            }
        }

        function doEmailReportsOk() {
            if (selectedReport == null || selectedReport === "") {
                showBinventoryMessageBox({
                    icon: "info",
                    message: "Please select a report.",
                    buttons: "ok"
                });
                return;
            }
            if (selectedReport === "Site Status Report") {
                openScheduleMaintenance(EMAIL_REPORT_EVENT_TYPE, { stack: true });
                return;
            }
            showBinventoryMessageBox({
                icon: "warn",
                message: "Invalid Report.",
                buttons: "ok"
            });
        }

        if (optEl) {
            optEl.addEventListener("click", function (e) {
                e.stopPropagation();
                selectSiteStatus();
            });
            optEl.addEventListener("dblclick", function (e) {
                e.stopPropagation();
                selectSiteStatus();
                doEmailReportsOk();
            });
        }
        if (lb) {
            lb.addEventListener("keydown", function (e) {
                if (e.key === "Enter") {
                    e.preventDefault();
                    doEmailReportsOk();
                }
            });
        }

        document.getElementById("erBtnSelect").addEventListener("click", function () {
            doEmailReportsOk();
        });

        clearSelection();
    }

    function simDatFileNameForDate(d) {
        d = d || new Date();
        var y = d.getFullYear();
        var m = String(d.getMonth() + 1).padStart(2, "0");
        var day = String(d.getDate()).padStart(2, "0");
        return "eBobWorkstationSystem_" + y + "-" + m + "-" + day + ".dat";
    }

    function simNormalizePath(p) {
        return String(p || "").replace(/\//g, "\\").trim();
    }

    function closeSimSystemExportImportModal() {
        var el = document.getElementById("backdropSimSysExportImport");
        if (el) {
            el.classList.remove("show");
            el.setAttribute("aria-hidden", "true");
        }
    }

    function closeSimSaveDatDialog() {
        var el = document.getElementById("backdropSimSaveDat");
        if (el) {
            el.classList.remove("show");
            el.setAttribute("aria-hidden", "true");
        }
    }

    function closeSimOpenDatDialog() {
        var el = document.getElementById("backdropSimOpenDat");
        if (el) {
            el.classList.remove("show");
            el.setAttribute("aria-hidden", "true");
        }
    }

    function openSimSystemExportImportModal() {
        var bd = document.getElementById("backdropSimSysExportImport");
        var cap = document.getElementById("simSysExpImpCapTitle");
        var lbl = document.getElementById("simSysLblTitle");
        var btn = document.getElementById("simSysBtnPrimary");
        var txt = document.getElementById("simSysTxtFileLocation");
        if (!bd || !cap || !lbl || !btn || !txt) return;
        var msgTitle = "Binventory Workstation";
        if (simSysExportImportMode === 1) {
            cap.textContent = "System Export - " + msgTitle;
            lbl.textContent = "System Export";
            btn.textContent = "Export";
            txt.value = simNormalizePath(SIM_FAKE_DOCS_ROOT + "\\" + simDatFileNameForDate());
        } else {
            cap.textContent = "System Import - " + msgTitle;
            lbl.textContent = "System Import";
            btn.textContent = "Import";
            txt.value = simNormalizePath(SIM_FAKE_DOCS_ROOT);
        }
        bd.classList.add("show");
        bd.setAttribute("aria-hidden", "false");
        txt.focus();
    }

    function exportData() {
        simSysExportImportMode = 1;
        openSimSystemExportImportModal();
    }

    function importData(file, done) {
        var r = new FileReader();
        r.onload = function () {
            try {
                var o = JSON.parse(r.result);
                if (!o.vessels || !Array.isArray(o.vessels)) throw new Error("Invalid file");
                state = Object.assign(state, o);
                saveState();
                refreshUI();
                toast("Configuration imported.");
                if (done) done(true);
            } catch (err) {
                toast("Import failed: invalid file.");
                if (done) done(false);
            }
        };
        r.readAsText(file);
    }

    function applySimImportedStateObject(o) {
        if (!o || !o.vessels || !Array.isArray(o.vessels)) return false;
        state = Object.assign(state, o);
        saveState();
        refreshUI();
        return true;
    }

    function doSimExportWrite(pathNorm) {
        var name = pathNorm.split("\\").pop();
        simSysExportSnapshotJson = JSON.stringify(state);
        var bytes = simSysExportSnapshotJson.length;
        var sizeStr = bytes < 1024 ? bytes + " bytes" : Math.ceil(bytes / 1024) + " KB";
        simSysExportDatMeta = {
            name: name,
            fullPath: pathNorm,
            dateModified: new Date().toLocaleString(),
            sizeStr: sizeStr,
            type: "DAT File"
        };
        closeSimSystemExportImportModal();
        showBinventoryMessageBox({
            title: "Binventory Workstation",
            icon: "info",
            message: "Export complete."
        });
    }

    function performSimSystemExport() {
        var txt = document.getElementById("simSysTxtFileLocation");
        if (!txt) return;
        var path = simNormalizePath(txt.value);
        if (!path || !/\.dat$/i.test(path)) {
            toast("Please enter a path ending in .dat");
            return;
        }
        var name = path.split("\\").pop();
        if (!/^eBobWorkstationSystem_\d{4}-\d{2}-\d{2}\.dat$/i.test(name)) {
            toast("Please use the eBob Workstation system naming pattern (eBobWorkstationSystem_YYYY-MM-DD.dat).");
            return;
        }
        if (simSysExportDatMeta && simSysExportDatMeta.fullPath === path) {
            showBinventoryMessageBox({
                title: "Binventory Workstation",
                icon: "warn",
                buttons: "okcancel",
                message: 'Do you want to overwrite the file named "' + path + '"?',
                onOk: function () {
                    doSimExportWrite(path);
                }
            });
            return;
        }
        doSimExportWrite(path);
    }

    function performSimSystemImport() {
        var txt = document.getElementById("simSysTxtFileLocation");
        if (!txt) return;
        var path = simNormalizePath(txt.value);
        if (!path) {
            showBinventoryMessageBox({
                title: "Binventory Workstation",
                icon: "warn",
                message: "Choose a file location or browse for the system data file."
            });
            return;
        }
        if (!simSysExportSnapshotJson || !simSysExportDatMeta) {
            showBinventoryMessageBox({
                title: "Binventory Workstation",
                icon: "warn",
                message: "There is no system export file in the simulated Documents folder. Run System Export first."
            });
            return;
        }
        var okPath =
            path === simSysExportDatMeta.fullPath ||
            path === simSysExportDatMeta.name ||
            path.endsWith("\\" + simSysExportDatMeta.name);
        if (!okPath) {
            showBinventoryMessageBox({
                title: "Binventory Workstation",
                icon: "warn",
                message: "The selected file does not match the system export in the simulated Documents folder."
            });
            return;
        }
        var o;
        try {
            o = JSON.parse(simSysExportSnapshotJson);
        } catch (e1) {
            toast("Import failed.");
            return;
        }
        if (!applySimImportedStateObject(o)) {
            toast("Import failed: invalid snapshot.");
            return;
        }
        closeSimSystemExportImportModal();
        showBinventoryMessageBox({
            title: "Binventory Workstation",
            icon: "warn",
            message:
                "Import complete.  The eBob 5 workstation and local services will now restart.\n\n" +
                "IMPORTANT:  If there are other workstations or services connected or networked to the eBob 5 database then they must also be restarted now.",
            onOk: function () {
                showUacForImportRestart();
            }
        });
    }

    /** Fills Open or Save Explorer file list (same simulated Documents folder). */
    function renderSimExplorerDatList(tbodyId) {
        var tb = document.getElementById(tbodyId);
        if (!tb) return;
        tb.innerHTML = "";
        if (!simSysExportDatMeta) {
            var tr0 = document.createElement("tr");
            tr0.innerHTML =
                '<td colspan="4" class="sim-open-dat-empty">This folder is empty.</td>';
            tb.appendChild(tr0);
            return;
        }
        var m = simSysExportDatMeta;
        var tr = document.createElement("tr");
        tr.className = "sim-open-dat-row";
        tr.dataset.fileName = m.name;
        var datIconSvg =
            '<svg viewBox="0 0 16 16" aria-hidden="true"><path fill="#e8e8e8" d="M3 1.5h5l1.5 1.5H13a1 1 0 011 1V14a1 1 0 01-1 1H3a1 1 0 01-1-1V2.5a1 1 0 011-1z"/><path fill="#6b9fde" d="M3 5.5h10v1H3v-1zm0 2.5h7v1H3V8zm0 2.5h5v1H3v-1z"/></svg>';
        tr.innerHTML =
            '<td>' +
            '<div class="ofi-name-cell">' +
            '<span class="ofi-file-ico">' +
            datIconSvg +
            "</span>" +
            '<span class="sim-open-dat-name">' +
            escapeHtml(m.name) +
            "</span>" +
            "</div></td>" +
            "<td>" +
            escapeHtml(m.dateModified) +
            "</td>" +
            "<td>" +
            escapeHtml(m.type) +
            "</td>" +
            "<td>" +
            escapeHtml(m.sizeStr) +
            "</td>";
        tb.appendChild(tr);
    }

    function renderSimOpenDatFileList() {
        renderSimExplorerDatList("simOpenDatTbody");
    }

    function renderSimSaveDatFileList() {
        renderSimExplorerDatList("simSaveDatTbody");
    }

    function wireSimSystemExportImport() {
        var bd = document.getElementById("backdropSimSysExportImport");
        var capClose = document.getElementById("simSysExpImpCapClose");
        var btnBrowse = document.getElementById("simSysBtnBrowse");
        var btnPrimary = document.getElementById("simSysBtnPrimary");
        var btnCancel = document.getElementById("simSysBtnCancel");
        if (!bd) return;
        if (capClose) {
            capClose.addEventListener("click", function () {
                closeSimSystemExportImportModal();
            });
        }
        if (btnCancel) {
            btnCancel.addEventListener("click", function () {
                closeSimSystemExportImportModal();
            });
        }
        if (btnBrowse) {
            btnBrowse.addEventListener("click", function () {
                if (simSysExportImportMode === 1) {
                    renderSimSaveDatFileList();
                    var fn = document.getElementById("simSaveDatFileName");
                    if (fn) fn.value = simDatFileNameForDate();
                    var bodSave = document.getElementById("simSaveDatTbody");
                    if (bodSave) {
                        bodSave.querySelectorAll(".sim-open-dat-row").forEach(function (r) {
                            r.classList.remove("sim-open-dat-row--sel");
                        });
                    }
                    var bsd = document.getElementById("backdropSimSaveDat");
                    if (bsd) {
                        bsd.classList.add("show");
                        bsd.setAttribute("aria-hidden", "false");
                    }
                } else {
                    renderSimOpenDatFileList();
                    var fn2 = document.getElementById("simOpenDatFileName");
                    var bod = document.getElementById("simOpenDatTbody");
                    if (fn2) fn2.value = "";
                    if (bod) {
                        bod.querySelectorAll(".sim-open-dat-row").forEach(function (r) {
                            r.classList.remove("sim-open-dat-row--sel");
                        });
                    }
                    var bodOpen = document.getElementById("backdropSimOpenDat");
                    if (bodOpen) {
                        bodOpen.classList.add("show");
                        bodOpen.setAttribute("aria-hidden", "false");
                    }
                }
            });
        }
        if (btnPrimary) {
            btnPrimary.addEventListener("click", function () {
                if (simSysExportImportMode === 1) {
                    performSimSystemExport();
                } else {
                    performSimSystemImport();
                }
            });
        }

        var saveBd = document.getElementById("backdropSimSaveDat");
        var saveClose = document.getElementById("simSaveDatCapClose");
        var saveCancel = document.getElementById("simSaveDatBtnCancel");
        var saveOk = document.getElementById("simSaveDatBtnSave");
        if (saveClose) {
            saveClose.addEventListener("click", closeSimSaveDatDialog);
        }
        if (saveCancel) {
            saveCancel.addEventListener("click", closeSimSaveDatDialog);
        }
        if (saveOk) {
            saveOk.addEventListener("click", function () {
                var fnEl = document.getElementById("simSaveDatFileName");
                var fn = (fnEl && fnEl.value) || simDatFileNameForDate();
                fn = fn.replace(/[/\\]/g, "").trim();
                if (!/.+\.dat$/i.test(fn)) fn = fn + ".dat";
                var txt = document.getElementById("simSysTxtFileLocation");
                if (txt) txt.value = simNormalizePath(SIM_FAKE_DOCS_ROOT + "\\" + fn);
                closeSimSaveDatDialog();
            });
        }
        if (saveBd) {
            saveBd.addEventListener("click", function (e) {
                if (e.target === saveBd) closeSimSaveDatDialog();
            });
        }

        var saveTb = document.getElementById("simSaveDatTbody");
        if (saveTb) {
            saveTb.addEventListener("click", function (e) {
                if (e.detail > 1) return;
                var row = e.target.closest(".sim-open-dat-row");
                if (!row) return;
                saveTb.querySelectorAll(".sim-open-dat-row").forEach(function (r) {
                    r.classList.remove("sim-open-dat-row--sel");
                });
                row.classList.add("sim-open-dat-row--sel");
                var fnS = document.getElementById("simSaveDatFileName");
                if (fnS) fnS.value = row.dataset.fileName || "";
            });
            saveTb.addEventListener("dblclick", function (e) {
                var row = e.target.closest(".sim-open-dat-row");
                if (row && saveOk) saveOk.click();
            });
        }

        var openBd = document.getElementById("backdropSimOpenDat");
        var openClose = document.getElementById("simOpenDatCapClose");
        var openCancel = document.getElementById("simOpenDatBtnCancel");
        var openOk = document.getElementById("simOpenDatBtnOpen");
        var openTb = document.getElementById("simOpenDatTbody");
        if (openClose) {
            openClose.addEventListener("click", closeSimOpenDatDialog);
        }
        if (openCancel) {
            openCancel.addEventListener("click", closeSimOpenDatDialog);
        }
        if (openTb) {
            openTb.addEventListener("click", function (e) {
                if (e.detail > 1) return;
                var row = e.target.closest(".sim-open-dat-row");
                if (!row) return;
                openTb.querySelectorAll(".sim-open-dat-row").forEach(function (r) {
                    r.classList.remove("sim-open-dat-row--sel");
                });
                row.classList.add("sim-open-dat-row--sel");
                var fn = document.getElementById("simOpenDatFileName");
                if (fn) fn.value = row.dataset.fileName || "";
            });
            openTb.addEventListener("dblclick", function (e) {
                var row = e.target.closest(".sim-open-dat-row");
                if (row && openOk) openOk.click();
            });
        }
        if (openOk) {
            openOk.addEventListener("click", function () {
                var fnEl = document.getElementById("simOpenDatFileName");
                var name = (fnEl && fnEl.value.trim()) || "";
                if (!simSysExportDatMeta || name !== simSysExportDatMeta.name) {
                    toast("Select the file that appears after System Export.");
                    return;
                }
                var txt = document.getElementById("simSysTxtFileLocation");
                if (txt) txt.value = simSysExportDatMeta.fullPath;
                closeSimOpenDatDialog();
            });
        }
        if (openBd) {
            openBd.addEventListener("click", function (e) {
                if (e.target === openBd) closeSimOpenDatDialog();
            });
        }

        document.querySelectorAll(".sim-open-dat-list-wrap").forEach(function (wrap) {
            wrap.addEventListener("click", function (e) {
                if (e.target.closest(".sim-open-dat-row")) return;
                var backdrop = wrap.closest("#backdropSimSaveDat, #backdropSimOpenDat");
                if (!backdrop) return;
                var tbody = wrap.querySelector("tbody");
                if (!tbody) return;
                tbody.querySelectorAll(".sim-open-dat-row").forEach(function (r) {
                    r.classList.remove("sim-open-dat-row--sel");
                });
                var isSave = backdrop.id === "backdropSimSaveDat";
                var fnEl = document.getElementById(isSave ? "simSaveDatFileName" : "simOpenDatFileName");
                if (fnEl) {
                    if (isSave) fnEl.value = simDatFileNameForDate();
                    else fnEl.value = "";
                }
            });
        });

        if (bd) {
            bd.addEventListener("click", function (e) {
                if (e.target === bd) closeSimSystemExportImportModal();
            });
        }
    }

    function dispatchAction(action) {
        switch (action) {
            case "logoff":
                performLogoff();
                break;
            case "login":
                performLoginFromMenu();
                break;
            case "change-password":
                openChangePassword();
                break;
            case "user-maintenance":
                openUserMaintenance();
                break;
            case "export-data":
                exportData();
                break;
            case "import-data":
                simSysExportImportMode = 2;
                openSimSystemExportImportModal();
                break;
            case "exit-app":
                exitToSimDesktop();
                break;
            case "vessel-maintenance":
                openVesselMaintenance();
                break;
            case "group-maintenance":
                openGroupMaintenance();
                break;
            case "schedule-maintenance":
                openScheduleMaintenance();
                break;
            case "temporary-group":
                openTemporaryGroup();
                break;
            case "measure-all":
                measureAll();
                break;
            case "email-reports":
                openEmailReports();
                break;
            case "contact-maintenance":
                openContactMaintenance();
                break;
            case "site-maintenance":
                openSiteMaintenance();
                break;
            case "sensor-networks":
                openSensorNetworks();
                break;
            case "email-setup":
                openEmailSetup();
                break;
            case "system-setup":
                openSystemSetup();
                break;
            case "operators-manual":
                document.getElementById("infoTitle").textContent = "Operator's Manual";
                document.getElementById("infoBody").textContent =
                    "Refer to the Binventory Workstation documentation for vessel setup, groups, schedules, specifications, and reports.";
                document.getElementById("backdropInfo").classList.add("show");
                break;
            default:
                break;
        }
    }

    /* —— init —— */
    (function bootstrapBinventory() {
        var pw = document.getElementById("pageWrap");
        if (pw) pw.classList.add("startup-sequence-active");
        loadState();
        runBinventoryStartupSequence(function () {
            refreshUI({ staggerVesselReveal: true });
            if (pendingImportRestartToast) {
                pendingImportRestartToast = false;
                toast("Binventory Workstation restarted after system import (simulation).");
            }
        });
    })();
    wireAssignContactsDialog();
    wireSimSystemExportImport();

    (function wireDesktopRelaunch() {
        var icon = document.getElementById("desktopEbobIcon");
        if (icon) {
            icon.addEventListener("click", function () {
                launchEbobFromDesktop();
            });
        }
        var tb = document.querySelector(".win-taskbar-ebob");
        if (tb) {
            tb.addEventListener("click", function () {
                var desk = document.getElementById("simDesktop");
                if (desk && !desk.hidden) {
                    launchEbobFromDesktop();
                }
            });
        }
    })();

    var menuRoots = document.querySelectorAll(".menu-root[data-menu]");
    /** True while a menu dropdown is open in "menu bar mode" (after first open) — hover switches top-level menus like Windows. */
    var menuBarOpen = false;

    function stripMenuOpenClasses() {
        menuRoots.forEach(function (m) {
            m.classList.remove("open");
        });
    }

    function closeMenus() {
        stripMenuOpenClasses();
        menuBarOpen = false;
    }

    menuRoots.forEach(function (root) {
        var btn = root.querySelector(":scope > button");
        if (!btn) return;
        root.addEventListener("mouseenter", function () {
            if (!menuBarOpen) return;
            stripMenuOpenClasses();
            root.classList.add("open");
        });
        btn.addEventListener("click", function (e) {
            e.stopPropagation();
            var wasOpen = root.classList.contains("open");
            stripMenuOpenClasses();
            if (wasOpen) {
                menuBarOpen = false;
            } else {
                root.classList.add("open");
                menuBarOpen = true;
            }
        });
    });
    document.addEventListener("click", closeMenus);
    menuRoots.forEach(function (root) {
        root.addEventListener("click", function (e) { e.stopPropagation(); });
    });
    document.querySelectorAll(".menu-sub-wrap").forEach(function (w) {
        w.addEventListener("click", function (e) {
            e.stopPropagation();
        });
    });

    var siteAssignmentMenuItems = document.getElementById("siteAssignmentMenuItems");
    if (siteAssignmentMenuItems) {
        siteAssignmentMenuItems.addEventListener("click", function (e) {
            var btn = e.target.closest(".sa-menu-item[data-site-id]");
            if (!btn) return;
            e.preventDefault();
            e.stopPropagation();
            var siteId = btn.getAttribute("data-site-id");
            if (!siteId || siteId === state.currentWorkstationSiteId) return;
            closeMenus();
            runSiteSwitchLoadingSequence(siteId, function () {
                var s = findSite(siteId);
                toast("Workstation reloaded — site: " + (s && s.name ? s.name : siteId) + " (simulation).");
            });
        });
    }
    var siteAssignmentBtn = document.querySelector(".site-assignment-root > button");
    if (siteAssignmentBtn) {
        siteAssignmentBtn.addEventListener("click", function () {
            refreshSiteAssignmentMenu();
        });
    }

    var bdInfo = document.getElementById("backdropInfo");
    var bdLogin = document.getElementById("backdropLogin");
    var bdPrint = document.getElementById("backdropPrint");
    var bdAbout = document.getElementById("backdropAbout");
    var bdReportSystemError = document.getElementById("backdropReportSystemError");
    var bdReportInventory = document.getElementById("backdropReportInventory");
    var bdReportHistory = document.getElementById("backdropReportHistory");

    var MS_MSG_TITLE = "Binventory Workstation";

    function dateToYMD(d) {
        var y = d.getFullYear();
        var m = String(d.getMonth() + 1).padStart(2, "0");
        var day = String(d.getDate()).padStart(2, "0");
        return y + "-" + m + "-" + day;
    }

    /** Display date like 3/18/2026 for report headers (from yyyy-mm-dd). */
    function formatDateForRptDisplay(isoYmd) {
        if (!isoYmd) return "";
        var d = new Date(isoYmd + "T12:00:00");
        if (isNaN(d.getTime())) return isoYmd;
        return d.toLocaleDateString(undefined, { month: "numeric", day: "numeric", year: "numeric" });
    }

    function defaultDateRangeFromNow() {
        var now = new Date();
        now.setHours(0, 0, 0, 0);
        var from = new Date(now);
        from.setDate(from.getDate() - 7);
        return { from: from, thru: now };
    }

    /** Tutorial validation: date range and length checks (same messages as the desktop app). */
    function validateDateRangeStrings(sFrom, sThru) {
        var sF = sFrom || "";
        var sT = sThru || "";
        if (sF.length > 25) return "From Date must be 25 characters or less.";
        if (sT.length > 25) return "Through Date must be 25 characters or less.";
        if (!sF || !sT) return "From Date is invalid.";
        var dFrom = new Date(sF + "T12:00:00");
        var dThru = new Date(sT + "T12:00:00");
        if (isNaN(dFrom.getTime())) return "From Date is invalid.";
        if (isNaN(dThru.getTime())) return "Through Date is invalid.";
        if (dFrom > dThru) {
            return "From Date cannot be more recent than Through Date and Through Date cannot be older than From Date.";
        }
        return null;
    }

    function vesselsForCurrentSiteSorted() {
        var sid = state.currentWorkstationSiteId;
        var list = state.vessels.filter(function (v) {
            return v.siteId == null || v.siteId === sid;
        });
        list.sort(function (a, b) {
            var ao = a.sortOrder != null ? a.sortOrder : 0;
            var bo = b.sortOrder != null ? b.sortOrder : 0;
            return ao - bo;
        });
        return list;
    }

    function vesselDisplayLabel(v) {
        return v.name || String(v.vesselNumericId != null ? v.vesselNumericId : "");
    }

    function populateInventoryVesselLists() {
        var avail = document.getElementById("invAvailList");
        var sel = document.getElementById("invSelList");
        if (!avail || !sel) return;
        while (avail.firstChild) avail.removeChild(avail.firstChild);
        while (sel.firstChild) sel.removeChild(sel.firstChild);
        vesselsForCurrentSiteSorted().forEach(function (v) {
            var o = document.createElement("option");
            o.value = v.id;
            o.textContent = vesselDisplayLabel(v);
            avail.appendChild(o);
        });
    }

    function updateInventoryMoveButtons() {
        var avail = document.getElementById("invAvailList");
        var sel = document.getElementById("invSelList");
        var b1 = document.getElementById("invBtnSel");
        var b2 = document.getElementById("invBtnSelAll");
        var b3 = document.getElementById("invBtnRem");
        var b4 = document.getElementById("invBtnRemAll");
        if (!avail || !sel) return;
        if (b1) b1.disabled = avail.selectedIndex < 0 || avail.options.length === 0;
        if (b2) b2.disabled = avail.options.length === 0;
        if (b3) b3.disabled = sel.selectedIndex < 0 || sel.options.length === 0;
        if (b4) b4.disabled = sel.options.length === 0;
    }

    function moveInventoryOption(fromId, toId, all) {
        var fromEl = document.getElementById(fromId);
        var toEl = document.getElementById(toId);
        if (!fromEl || !toEl) return;
        var toMove = [];
        if (all) {
            var i;
            for (i = 0; i < fromEl.options.length; i++) toMove.push(fromEl.options[i]);
        } else {
            if (fromEl.selectedIndex < 0) return;
            toMove.push(fromEl.options[fromEl.selectedIndex]);
        }
        toMove.forEach(function (o) {
            toEl.appendChild(o);
        });
        updateInventoryMoveButtons();
    }

    function populateHistoryVesselList() {
        var sel = document.getElementById("histVesselList");
        if (!sel) return;
        while (sel.firstChild) sel.removeChild(sel.firstChild);
        vesselsForCurrentSiteSorted().forEach(function (v) {
            var o = document.createElement("option");
            o.value = v.id;
            o.textContent = vesselDisplayLabel(v);
            sel.appendChild(o);
        });
        if (sel.options.length) sel.selectedIndex = 0;
    }

    /** Tutorial: dates enabled only for Comma Delimited export. */
    function applyInventoryOutputTypeUI() {
        var csv = document.getElementById("invOutCsv");
        var row = document.getElementById("invDateRow");
        var from = document.getElementById("invFromDate");
        var thru = document.getElementById("invThruDate");
        if (!csv || !row || !from || !thru) return;
        var csvEnabled = csv.checked;
        from.disabled = !csvEnabled;
        thru.disabled = !csvEnabled;
        row.querySelectorAll("label").forEach(function (lbl) {
            lbl.style.opacity = csvEnabled ? "1" : "0.55";
        });
    }

    /** Tutorial: dates enabled only when “Use Date Range” is selected. */
    function applyHistoryRadioUI() {
        var radDates = document.getElementById("histRadDates");
        var row = document.getElementById("histDateRow");
        var from = document.getElementById("histFromDate");
        var thru = document.getElementById("histThruDate");
        if (!radDates || !row || !from || !thru) return;
        var enable = radDates.checked;
        from.disabled = !enable;
        thru.disabled = !enable;
        row.querySelectorAll("label").forEach(function (lbl) {
            lbl.style.opacity = enable ? "1" : "0.55";
        });
    }

    function mockReportSiteLabel() {
        var s = findSite(state.currentWorkstationSiteId);
        return s && s.name ? s.name : "Current site";
    }

    function mockRng(seed) {
        var x = Math.sin(seed) * 10000;
        return x - Math.floor(x);
    }

    /**
     * Crystal-style report viewer shell: toolbar, gray canvas, white page, three-part header, status bar.
     */
    function buildCrystalViewerChrome(opts, innerBodyHtml) {
        var now = new Date();
        var fromStr = opts.fromStr != null ? opts.fromStr : "";
        var thruStr = opts.thruStr != null ? opts.thruStr : "";
        var reportHeading = opts.reportHeading || "Report";
        var timeStr = now.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit", second: "2-digit" });
        var dateStr = formatDateForRptDisplay(dateToYMD(now));
        var toolbar =
            '<div class="rpt-crystal-toolbar">' +
            '<div class="rpt-crystal-toolbar-btns">' +
            '<span class="rpt-crystal-tb-btn">Export</span>' +
            '<span class="rpt-crystal-tb-btn">Print</span>' +
            '<span class="rpt-crystal-tb-btn">Refresh</span>' +
            '<span class="rpt-crystal-tb-sep">|</span>' +
            '<span class="rpt-crystal-tb-btn">Find</span>' +
            '<span class="rpt-crystal-tb-sep">|</span>' +
            '<span class="rpt-crystal-tb-btn">&lt;&lt;</span>' +
            '<span class="rpt-crystal-tb-btn">&lt;</span>' +
            '<span style="padding:0 6px">Page 1 of 1</span>' +
            '<span class="rpt-crystal-tb-btn">&gt;</span>' +
            '<span class="rpt-crystal-tb-btn">&gt;&gt;</span>' +
            "</div>" +
            '<div class="rpt-crystal-toolbar-brand">Report Viewer</div></div>' +
            '<div class="rpt-crystal-tabrow"><span class="rpt-crystal-tab-active">Main Report</span></div>';
        var header =
            '<div class="rpt-crystal-header-grid">' +
            '<div class="rpt-crystal-h-left">From Date: ' +
            escapeHtml(fromStr) +
            "<br>Thru Date: " +
            escapeHtml(thruStr) +
            "</div>" +
            '<div class="rpt-crystal-h-center"><strong>Binventory Workstation</strong><br><strong>' +
            escapeHtml(reportHeading) +
            "</strong></div>" +
            '<div class="rpt-crystal-h-right">' +
            escapeHtml(dateStr) +
            "<br>" +
            escapeHtml(timeStr) +
            "</div></div>" +
            '<hr class="rpt-crystal-header-line" />';
        return (
            '<div class="rpt-crystal-shell">' +
            toolbar +
            '<div class="rpt-crystal-canvas">' +
            '<div class="rpt-crystal-page">' +
            header +
            innerBodyHtml +
            '<div class="rpt-crystal-table-end"></div></div></div>' +
            '<div class="rpt-crystal-statusbar">Current Page No.: 1 &nbsp;&nbsp; Total Page No.: 1 &nbsp;&nbsp; Zoom Factor: 100%</div></div>'
        );
    }

    function buildMockSystemErrorReportHtml(fromYmd, thruYmd) {
        var fromDisp = formatDateForRptDisplay(fromYmd);
        var thruDisp = formatDateForRptDisplay(thruYmd);
        var d0 = new Date(fromYmd + "T12:00:00");
        var d1 = new Date(thruYmd + "T12:00:00");
        var ms = Math.max(0, d1.getTime() - d0.getTime());
        var mock = [
            {
                desc: "Object reference not set to an instance of an object.",
                stack:
                    "Stack Trace: at Tutorial.MockSensorNetwork..ctor(NetworkRecord rec) in MockSensorNetwork.vb:line 332",
                user: "SYSTEM",
                sys: "BobsBO",
                sysVer: "5.4.9.19851",
                dllVer: "5.4.9",
                dbVer: "4.38"
            },
            {
                desc: "Timeout waiting for sensor response.",
                stack:
                    "Stack Trace: at Tutorial.MockSensorNetwork.PollMeasurement() in MockSensorNetwork.vb:line 118",
                user: "DemoUser",
                sys: "BobsBO",
                sysVer: "5.4.9.19851",
                dllVer: "5.4.9",
                dbVer: "4.38"
            },
            {
                desc: "Unable to connect to network service.",
                stack: "Stack Trace: at Tutorial.MockNetworkClient.Connect() in MockNetworkClient.vb:line 45",
                user: "SYSTEM",
                sys: "BobsBO",
                sysVer: "5.4.9.19851",
                dllVer: "5.4.9",
                dbVer: "4.38"
            },
            {
                desc: "Invalid calibration value.",
                stack: "Stack Trace: at Tutorial.MockCalibration.Apply() in MockCalibration.vb:line 201",
                user: "admin",
                sys: "BobsBO",
                sysVer: "5.4.9.19851",
                dllVer: "5.4.9",
                dbVer: "4.38"
            }
        ];
        var rows = [];
        var i;
        for (i = 0; i < mock.length; i++) {
            var t = mockRng(i + ms + 1) * (ms || 86400000);
            var dt = new Date(d0.getTime() + t);
            var errDate = dt.toLocaleString(undefined, {
                month: "numeric",
                day: "numeric",
                year: "numeric",
                hour: "numeric",
                minute: "2-digit",
                second: "2-digit",
                hour12: true
            });
            var m = mock[i];
            rows.push(
                "<tr><td>" +
                    escapeHtml(errDate) +
                    '</td><td class="rpt-crystal-desc-cell"><div class="rpt-crystal-err-text">' +
                    escapeHtml(m.desc) +
                    '</div><div class="rpt-crystal-stack">' +
                    escapeHtml(m.stack) +
                    "</div></td><td>" +
                    escapeHtml(m.user) +
                    "</td><td>" +
                    escapeHtml(m.sys) +
                    "</td><td>" +
                    escapeHtml(m.sysVer) +
                    "</td><td>" +
                    escapeHtml(m.dllVer) +
                    "</td><td>" +
                    escapeHtml(m.dbVer) +
                    "</td></tr>"
            );
        }
        var table =
            '<table class="rpt-crystal-data-table"><thead><tr>' +
            "<th>Error Date</th><th>Error Description</th><th>User ID</th><th>System Name</th><th>System Ver.</th><th>DLL Ver.</th><th>DB Ver.</th>" +
            "</tr></thead><tbody>" +
            rows.join("") +
            "</tbody></table>";
        return buildCrystalViewerChrome(
            { reportHeading: "System Error Report", fromStr: fromDisp, thruStr: thruDisp },
            table
        );
    }

    function sortVesselsForReport(ids, sortIdx) {
        var vessels = ids.map(function (id) {
            return findVessel(id);
        }).filter(Boolean);
        if (sortIdx === 1) {
            vessels.sort(function (a, b) {
                return String(a.name || "").localeCompare(String(b.name || ""));
            });
        } else if (sortIdx === 2) {
            vessels.sort(function (a, b) {
                return String(a.contents || a.product || "").localeCompare(String(b.contents || b.product || ""));
            });
        }
        return vessels;
    }

    function buildMockInventoryTableHtml(vesselIds, sortIdx, fromYmd, thruYmd, reportHeading) {
        var fromDisp = formatDateForRptDisplay(fromYmd);
        var thruDisp = formatDateForRptDisplay(thruYmd);
        var vessels = sortVesselsForReport(vesselIds, sortIdx);
        var rows = vessels.map(function (v) {
            return (
                "<tr><td>" +
                escapeHtml(v.name) +
                "</td><td>" +
                escapeHtml(String(v.pctFull != null ? v.pctFull : "")) +
                "%</td><td>" +
                escapeHtml(String(v.volumeCuFt != null ? v.volumeCuFt : "")) +
                "</td><td>" +
                escapeHtml(String(v.weightLb != null ? v.weightLb : "")) +
                "</td><td>" +
                escapeHtml(v.status || "") +
                "</td></tr>"
            );
        });
        var table =
            '<table class="rpt-crystal-data-table rpt-inv-table"><thead><tr>' +
            "<th>Vessel</th><th>% Full</th><th>Volume</th><th>Weight</th><th>Status</th>" +
            "</tr></thead><tbody>" +
            rows.join("") +
            "</tbody></table>";
        return buildCrystalViewerChrome(
            {
                reportHeading: reportHeading || "Inventory Report",
                fromStr: fromDisp,
                thruStr: thruDisp
            },
            table
        );
    }

    function buildMockInventoryBarHtml(vesselIds, sortIdx, fromYmd, thruYmd) {
        var fromDisp = formatDateForRptDisplay(fromYmd);
        var thruDisp = formatDateForRptDisplay(thruYmd);
        var vessels = sortVesselsForReport(vesselIds, sortIdx);
        var maxPct = 1;
        vessels.forEach(function (v) {
            var p = parseFloat(v.pctFull);
            if (!isNaN(p) && p > maxPct) maxPct = p;
        });
        var bars = vessels.map(function (v) {
            var p = parseFloat(v.pctFull);
            if (isNaN(p)) p = 0;
            var w = maxPct > 0 ? Math.round((p / maxPct) * 100) : 0;
            return (
                '<div class="rpt-mock-bar-row"><span class="rpt-mock-bar-name">' +
                escapeHtml(v.name) +
                '</span><div class="rpt-mock-bar-track"><div class="rpt-mock-bar-fill" style="width:' +
                w +
                '%"></div></div><span>' +
                String(p) +
                "%</span></div>"
            );
        });
        var inner =
            '<p class="rpt-crystal-chart-caption">Fill level by vessel (mock chart)</p>' + bars.join("");
        return buildCrystalViewerChrome(
            { reportHeading: "Inventory Report — Bar graph", fromStr: fromDisp, thruStr: thruDisp },
            inner
        );
    }

    function buildMockInventoryCsv(vesselIds, fromYmd, thruYmd) {
        var lines = ["VesselName,PercentFull,Volume,Weight,MeasurementDate"];
        vesselIds.forEach(function (id) {
            var v = findVessel(id);
            if (!v) return;
            var row = [
                v.name,
                v.pctFull,
                v.volumeCuFt,
                v.weightLb,
                thruYmd || ""
            ].map(function (c) {
                return '"' + String(c).replace(/"/g, '""') + '"';
            });
            lines.push(row.join(","));
        });
        return lines.join("\r\n");
    }

    function sampleValueForVessel(v, yAxis) {
        if (!v) return 0;
        var base = parseFloat(v.volumeCuFt);
        if (isNaN(base)) base = 1000;
        var pct = parseFloat(v.pctFull);
        if (isNaN(pct)) pct = 50;
        var h = parseFloat(v.heightFt);
        if (isNaN(h)) h = 0;
        var cap = parseFloat(v.capacityHeightFt);
        if (isNaN(cap)) cap = 14;
        var wlb = parseFloat(v.weightLb);
        if (isNaN(wlb)) wlb = 0;
        switch (yAxis) {
            case 0:
                return Math.round(base * (pct / 100));
            case 1:
                return Math.round(wlb);
            case 2:
                return h;
            case 3:
                return Math.round(base * ((100 - pct) / 100));
            case 4:
                return Math.round(wlb * 0.3);
            case 5:
                return Math.max(0, cap - h);
            default:
                return 0;
        }
    }

    function buildMockHistoryChartHtml(vesselId, radioOpt, yAxis, fromYmd, thruYmd) {
        var v = findVessel(vesselId);
        if (!v) return "<p>No vessel data.</p>";
        var n = radioOpt === 1 ? 30 : radioOpt === 2 ? 60 : 14;
        var d0 = new Date(fromYmd + "T12:00:00");
        var d1 = new Date(thruYmd + "T12:00:00");
        if (radioOpt === 3) {
            n = Math.max(2, Math.ceil((d1 - d0) / 86400000) + 1);
            if (n > 60) n = 60;
        }
        var yLabels = [
            "Product Volume",
            "Product Weight",
            "Product Height",
            "Headroom Volume",
            "Headroom Weight",
            "Headroom Height"
        ];
        var yAxisTitle = yLabels[yAxis] || "Value";
        var base = sampleValueForVessel(v, yAxis);
        var pts = [];
        var i;
        for (i = 0; i < n; i++) {
            var jitter = (mockRng(i + n + base) - 0.5) * 0.2 * (base || 1);
            pts.push(Math.max(0, base + jitter));
        }
        var w = 480;
        var h = 200;
        var pad = 28;
        var maxY = Math.max.apply(null, pts);
        var minY = Math.min.apply(null, pts);
        if (maxY === minY) {
            minY = 0;
            maxY = maxY + 1;
        }
        var span = n - 1 || 1;
        var yspan = maxY - minY || 1;
        var pointStr = pts
            .map(function (p, idx) {
                var px = pad + (idx / span) * (w - 2 * pad);
                var py = pad + (1 - (p - minY) / yspan) * (h - 2 * pad);
                return Math.round(px) + "," + Math.round(py);
            })
            .join(" ");
        var svg =
            '<svg class="rpt-mock-chart-svg" viewBox="0 0 ' +
            w +
            " " +
            h +
            '" xmlns="http://www.w3.org/2000/svg"><rect fill="#fafafa" width="' +
            w +
            '" height="' +
            h +
            '"/><polyline fill="none" stroke="#1a4d8d" stroke-width="2" points="' +
            pointStr +
            '"/></svg>';
        var fromDisp = formatDateForRptDisplay(fromYmd);
        var thruDisp = formatDateForRptDisplay(thruYmd);
        var inner =
            '<p class="rpt-crystal-chart-caption"><strong>Vessel:</strong> ' +
            escapeHtml(v.name) +
            " &nbsp;|&nbsp; <strong>Y-axis:</strong> " +
            escapeHtml(yAxisTitle) +
            "</p>" +
            svg;
        return buildCrystalViewerChrome(
            { reportHeading: "Inventory History Report", fromStr: fromDisp, thruStr: thruDisp },
            inner
        );
    }

    function openMockReportModal(title, bodyHtml) {
        var footer =
            '<button type="button" data-mock-report-print>Print…</button><button type="button" class="secondary" data-close-app>Close</button>';
        openAppModal(title, bodyHtml, footer, "modal-report-preview modal-report-crystal");
    }

    function openSystemErrorReportModal() {
        var h = document.getElementById("serHeading");
        if (h) h.textContent = "System Error Report - " + MS_MSG_TITLE;
        var dr = defaultDateRangeFromNow();
        var fromEl = document.getElementById("serFromDate");
        var thruEl = document.getElementById("serThruDate");
        if (fromEl) fromEl.value = dateToYMD(dr.from);
        if (thruEl) thruEl.value = dateToYMD(dr.thru);
        if (bdReportSystemError) bdReportSystemError.classList.add("show");
    }

    function openInventoryReportModal() {
        var h = document.getElementById("invHeading");
        if (h) h.textContent = "Inventory Report - " + MS_MSG_TITLE;
        var rep = document.getElementById("invOutReport");
        if (rep) rep.checked = true;
        var dr = defaultDateRangeFromNow();
        var fromEl = document.getElementById("invFromDate");
        var thruEl = document.getElementById("invThruDate");
        if (fromEl) fromEl.value = dateToYMD(dr.from);
        if (thruEl) thruEl.value = dateToYMD(dr.thru);
        var sort = document.getElementById("invSortOrder");
        if (sort) sort.selectedIndex = 0;
        populateInventoryVesselLists();
        applyInventoryOutputTypeUI();
        updateInventoryMoveButtons();
        if (bdReportInventory) bdReportInventory.classList.add("show");
    }

    function openHistoryReportModal() {
        var h = document.getElementById("histHeading");
        if (h) h.textContent = "Inventory Report - " + MS_MSG_TITLE;
        var r30 = document.getElementById("histRad30");
        if (r30) r30.checked = true;
        var dr = defaultDateRangeFromNow();
        var fromEl = document.getElementById("histFromDate");
        var thruEl = document.getElementById("histThruDate");
        if (fromEl) fromEl.value = dateToYMD(dr.from);
        if (thruEl) thruEl.value = dateToYMD(dr.thru);
        var yax = document.getElementById("histChartYAxis");
        if (yax) yax.selectedIndex = 0;
        populateHistoryVesselList();
        applyHistoryRadioUI();
        if (bdReportHistory) bdReportHistory.classList.add("show");
    }

    ["invOutReport", "invOutBar", "invOutCsv"].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.addEventListener("change", applyInventoryOutputTypeUI);
    });
    ["histRad30", "histRad60", "histRadDates"].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.addEventListener("change", applyHistoryRadioUI);
    });

    var invBtnSel = document.getElementById("invBtnSel");
    if (invBtnSel) invBtnSel.addEventListener("click", function () { moveInventoryOption("invAvailList", "invSelList", false); });
    var invBtnSelAll = document.getElementById("invBtnSelAll");
    if (invBtnSelAll) invBtnSelAll.addEventListener("click", function () { moveInventoryOption("invAvailList", "invSelList", true); });
    var invBtnRem = document.getElementById("invBtnRem");
    if (invBtnRem) invBtnRem.addEventListener("click", function () { moveInventoryOption("invSelList", "invAvailList", false); });
    var invBtnRemAll = document.getElementById("invBtnRemAll");
    if (invBtnRemAll) invBtnRemAll.addEventListener("click", function () { moveInventoryOption("invSelList", "invAvailList", true); });

    var invAvailList = document.getElementById("invAvailList");
    var invSelList = document.getElementById("invSelList");
    if (invAvailList) {
        invAvailList.addEventListener("change", updateInventoryMoveButtons);
        invAvailList.addEventListener("dblclick", function () { moveInventoryOption("invAvailList", "invSelList", false); });
    }
    if (invSelList) {
        invSelList.addEventListener("change", updateInventoryMoveButtons);
        invSelList.addEventListener("dblclick", function () { moveInventoryOption("invSelList", "invAvailList", false); });
    }

    function closeReportSystemErrorBackdrop() {
        if (bdReportSystemError) bdReportSystemError.classList.remove("show");
    }
    function closeReportInventoryBackdrop() {
        if (bdReportInventory) bdReportInventory.classList.remove("show");
    }
    function closeReportHistoryBackdrop() {
        if (bdReportHistory) bdReportHistory.classList.remove("show");
    }

    var serClose = document.getElementById("serClose");
    if (serClose) serClose.addEventListener("click", closeReportSystemErrorBackdrop);
    var serCapClose = document.getElementById("serCapClose");
    if (serCapClose) serCapClose.addEventListener("click", closeReportSystemErrorBackdrop);
    var invClose = document.getElementById("invClose");
    if (invClose) invClose.addEventListener("click", closeReportInventoryBackdrop);
    var invCapClose = document.getElementById("invCapClose");
    if (invCapClose) invCapClose.addEventListener("click", closeReportInventoryBackdrop);
    var histClose = document.getElementById("histClose");
    if (histClose) histClose.addEventListener("click", closeReportHistoryBackdrop);
    var histCapClose = document.getElementById("histCapClose");
    if (histCapClose) histCapClose.addEventListener("click", closeReportHistoryBackdrop);

    var serCreate = document.getElementById("serCreate");
    if (serCreate) {
        serCreate.addEventListener("click", function () {
            var fromEl = document.getElementById("serFromDate");
            var thruEl = document.getElementById("serThruDate");
            var err = validateDateRangeStrings(fromEl && fromEl.value, thruEl && thruEl.value);
            if (err) {
                toast(err);
                return;
            }
            if (bdReportSystemError) bdReportSystemError.classList.remove("show");
            openMockReportModal(
                "System Error Report",
                buildMockSystemErrorReportHtml(fromEl && fromEl.value, thruEl && thruEl.value)
            );
        });
    }

    var invCreate = document.getElementById("invCreate");
    if (invCreate) {
        invCreate.addEventListener("click", function () {
            var outCsv = document.getElementById("invOutCsv");
            var outBar = document.getElementById("invOutBar");
            var outType = outCsv && outCsv.checked ? "2" : outBar && outBar.checked ? "3" : "1";
            var fromEl = document.getElementById("invFromDate");
            var thruEl = document.getElementById("invThruDate");
            var sel = document.getElementById("invSelList");
            if (!sel || sel.options.length === 0) {
                toast("At least one Vessel must be selected.");
                return;
            }
            if (outType === "2") {
                var err = validateDateRangeStrings(fromEl && fromEl.value, thruEl && thruEl.value);
                if (err) {
                    toast(err);
                    return;
                }
            }
            var sortEl = document.getElementById("invSortOrder");
            var sortIdx = sortEl ? parseInt(sortEl.value, 10) : 0;
            var ids = [];
            var i;
            for (i = 0; i < sel.options.length; i++) ids.push(sel.options[i].value);
            if (bdReportInventory) bdReportInventory.classList.remove("show");
            if (outType === "2") {
                var csv = buildMockInventoryCsv(ids, fromEl && fromEl.value, thruEl && thruEl.value);
                downloadTextFile("inventory_report.csv", csv, "text/csv;charset=utf-8");
                toast("CSV downloaded (mock data).");
                openMockReportModal(
                    "Inventory Report (preview)",
                    buildMockInventoryTableHtml(
                        ids,
                        sortIdx,
                        fromEl && fromEl.value,
                        thruEl && thruEl.value,
                        "Inventory Report (preview)"
                    )
                );
            } else if (outType === "3") {
                openMockReportModal(
                    "Inventory Report — Bar graph",
                    buildMockInventoryBarHtml(ids, sortIdx, fromEl && fromEl.value, thruEl && thruEl.value)
                );
            } else {
                openMockReportModal(
                    "Inventory Report",
                    buildMockInventoryTableHtml(
                        ids,
                        sortIdx,
                        fromEl && fromEl.value,
                        thruEl && thruEl.value,
                        "Inventory Report"
                    )
                );
            }
        });
    }

    var histCreate = document.getElementById("histCreate");
    if (histCreate) {
        histCreate.addEventListener("click", function () {
            var radDates = document.getElementById("histRadDates");
            var fromEl = document.getElementById("histFromDate");
            var thruEl = document.getElementById("histThruDate");
            var vsel = document.getElementById("histVesselList");
            if (!vsel || vsel.selectedIndex < 0 || vsel.options.length === 0) {
                toast("Select a Vessel.");
                return;
            }
            if (radDates && radDates.checked) {
                var err = validateDateRangeStrings(fromEl && fromEl.value, thruEl && thruEl.value);
                if (err) {
                    toast(err);
                    return;
                }
            }
            var h30 = document.getElementById("histRad30");
            var h60 = document.getElementById("histRad60");
            var radioOpt = h30 && h30.checked ? 1 : h60 && h60.checked ? 2 : 3;
            var yAxisEl = document.getElementById("histChartYAxis");
            var yAxis = yAxisEl ? parseInt(yAxisEl.value, 10) : 0;
            var vid = vsel.options[vsel.selectedIndex].value;
            if (bdReportHistory) bdReportHistory.classList.remove("show");
            openMockReportModal(
                "Inventory History Report",
                buildMockHistoryChartHtml(vid, radioOpt, yAxis, fromEl && fromEl.value, thruEl && thruEl.value)
            );
        });
    }

    document.querySelectorAll("#backdropInfo [data-close=info]").forEach(function (b) {
        b.addEventListener("click", function () {
            bdInfo.classList.remove("show");
        });
    });

    document.querySelectorAll("[data-close=vsimport]").forEach(function (b) {
        b.addEventListener("click", function () {
            var bdVs = document.getElementById("backdropVsImportPaste");
            if (bdVs) bdVs.classList.remove("show");
        });
    });
    var vsImportPasteOk = document.getElementById("vsImportPasteOk");
    if (vsImportPasteOk) {
        vsImportPasteOk.addEventListener("click", function () {
            var taEl = document.getElementById("vsImportPasteTa");
            var rawVal = taEl ? taEl.value : "";
            var bdVs = document.getElementById("backdropVsImportPaste");
            if (bdVs) bdVs.classList.remove("show");
            var parsed = parseVsCustomImportText(rawVal);
            if (parsed.length < 2) {
                showBinventoryMessageBox({
                    icon: "warn",
                    message:
                        "Either the clipboard was empty or it did not contain a table in a supported format.",
                    buttons: "ok"
                });
                return;
            }
            applyVsCustomParsedRowsToTable(parsed);
            showBinventoryMessageBox({
                icon: "info",
                message: "The custom/lookup table was imported successfully.",
                buttons: "ok"
            });
        });
    }

    document.getElementById("btnLoginOk").addEventListener("click", function () {
        var uidInput = document.getElementById("uid");
        var raw = uidInput ? uidInput.value.trim() : "";
        state.currentUser = raw || "User";
        var match = state.users.filter(function (x) {
            return x.name === state.currentUser || String(x.userId) === state.currentUser;
        })[0];
        if (match) {
            state.currentUser = match.name;
            match.lastLogon = new Date().toLocaleString();
        }
        saveState();
        bdLogin.classList.remove("show");
        bdLogin.setAttribute("aria-hidden", "true");
        syncMenuStripForSession();
        refreshUI({ staggerVesselReveal: true });
        updateTitleBar();
        toast("Logged in as " + state.currentUser);
    });
    document.querySelectorAll("[data-close=login]").forEach(function (b) {
        b.addEventListener("click", function () {
            bdLogin.classList.remove("show");
            bdLogin.setAttribute("aria-hidden", "true");
        });
    });

    document.querySelectorAll("#backdropPrint [data-close=print]").forEach(function (b) {
        b.addEventListener("click", function () {
            bdPrint.classList.remove("show");
        });
    });
    document.getElementById("btnPrintSelect").addEventListener("click", function () {
        var rs = document.getElementById("reportSelect");
        if (!rs || rs.selectedIndex < 0) {
            showBinventoryMessageBox({
                icon: "info",
                message: "Please select a report.",
                buttons: "ok"
            });
            return;
        }
        var val = rs.value;
        if (val === "system") openSystemErrorReportModal();
        else if (val === "inventory") openInventoryReportModal();
        else if (val === "history") openHistoryReportModal();
        else {
            showBinventoryMessageBox({
                icon: "warn",
                message: "Invalid report selection.",
                buttons: "ok"
            });
        }
    });

    document.querySelectorAll("[data-close=about]").forEach(function (b) {
        b.addEventListener("click", function () { bdAbout.classList.remove("show"); });
    });

    document.getElementById("menuStrip").addEventListener(
        "click",
        function (e) {
            var mlogin = e.target.closest("button[data-modal=login]");
            if (mlogin) {
                e.preventDefault();
                e.stopPropagation();
                closeMenus();
                bdLogin.classList.add("show");
                bdLogin.setAttribute("aria-hidden", "false");
                return;
            }
            var mprint = e.target.closest("button[data-modal=print-reports]");
            if (mprint) {
                e.preventDefault();
                e.stopPropagation();
                closeMenus();
                bdPrint.classList.add("show");
                return;
            }
            var mabout = e.target.closest("button[data-modal=about]");
            if (mabout) {
                e.preventDefault();
                e.stopPropagation();
                closeMenus();
                bdAbout.classList.add("show");
                return;
            }
            var actBtn = e.target.closest("button[data-action]");
            if (actBtn) {
                e.preventDefault();
                e.stopPropagation();
                closeMenus();
                dispatchAction(actBtn.getAttribute("data-action"));
                return;
            }
        },
        true
    );

    document.querySelector(".title-btns").addEventListener("click", function (e) {
        var b = e.target.closest("button[data-sim]");
        if (!b) return;
        var a = b.getAttribute("data-sim");
        if (a === "close") exitToSimDesktop();
        else toast("Window: " + a + " (simulation).");
    });

    document.getElementById("appShell").addEventListener("click", function (e) {
        var a = e.target.closest("[data-vessel-action]");
        if (!a) return;
        var act = a.getAttribute("data-vessel-action");
        var card = a.closest(".vessel");
        var id = card ? card.dataset.vesselId : null;
        var v = id ? findVessel(id) : null;
        if (act === "measure" && v) measureVessel(v);
        else if (act === "multi") toast("Multi-vessel options (simulation).");
    });

    document.getElementById("appShell").addEventListener("change", function (e) {
        if (e.target.matches('[data-vessel-action="headroom"]')) {
            var card = e.target.closest(".vessel");
            var v = card ? findVessel(card.dataset.vesselId) : null;
            if (v) {
                v.headroom = e.target.checked;
                saveState();
                renderGrid();
            }
        }
    });

    document.addEventListener("keydown", function (e) {
        if (e.key === "F1") {
            e.preventDefault();
            dispatchAction("operators-manual");
            return;
        }
        if (e.key === "Escape") {
            var bOpenDat = document.getElementById("backdropSimOpenDat");
            if (bOpenDat && bOpenDat.classList.contains("show")) {
                e.preventDefault();
                closeSimOpenDatDialog();
                return;
            }
            var bSaveDat = document.getElementById("backdropSimSaveDat");
            if (bSaveDat && bSaveDat.classList.contains("show")) {
                e.preventDefault();
                closeSimSaveDatDialog();
                return;
            }
            var bSysExp = document.getElementById("backdropSimSysExportImport");
            if (bSysExp && bSysExp.classList.contains("show")) {
                e.preventDefault();
                closeSimSystemExportImportModal();
                return;
            }
            var winStartMenu = document.getElementById("winStartMenu");
            var winStartBackdrop = document.getElementById("winStartBackdrop");
            var winStartBtn = document.getElementById("winStartBtn");
            if (winStartMenu && !winStartMenu.hidden && winStartBackdrop && winStartBtn) {
                var wss = document.getElementById("winStartSearch");
                if (wss) wss.value = "";
                winStartMenu.classList.remove("win-start-menu--search");
                var wFly = document.getElementById("winSearchFlyout");
                var wHome = document.getElementById("winStartHomeContent");
                if (wFly) wFly.hidden = true;
                if (wHome) wHome.hidden = false;
                winStartBackdrop.hidden = true;
                winStartBackdrop.setAttribute("aria-hidden", "true");
                winStartMenu.hidden = true;
                winStartBtn.setAttribute("aria-expanded", "false");
                e.preventDefault();
                return;
            }
            if (document.getElementById("backdropSnSetup").classList.contains("show")) {
                closeSensorNetworkSetup();
                return;
            }
            var bacAssign = document.getElementById("backdropAssignContacts");
            if (bacAssign && bacAssign.classList.contains("show")) {
                e.preventDefault();
                closeAssignContactsDialog();
                return;
            }
            var rfTop = document.querySelector(".backdrop-report-filter.show");
            if (rfTop) {
                rfTop.classList.remove("show");
                e.preventDefault();
                return;
            }
            closeMenus();
            bdInfo.classList.remove("show");
            bdLogin.classList.remove("show");
            bdPrint.classList.remove("show");
            bdAbout.classList.remove("show");
            if (bdApp.classList.contains("show") && appModalShell.classList.contains("modal-vessel-setup") && vmSetupEditingId) {
                e.preventDefault();
                closeVesselSetupDiscardChanges();
            } else {
                closeAppModal();
            }
        }
        if (e.target.matches("input, textarea, select")) return;
        if (e.ctrlKey && e.key === "v") {
            e.preventDefault();
            dispatchAction("vessel-maintenance");
        }
        if (e.ctrlKey && e.key === "g") {
            e.preventDefault();
            dispatchAction("group-maintenance");
        }
        if (e.ctrlKey && e.key === "s") {
            e.preventDefault();
            dispatchAction("schedule-maintenance");
        }
        if (e.key === "F5") {
            var mMeas = document.querySelector('button[data-action="measure-all"]');
            if (mMeas && mMeas.disabled) return;
            e.preventDefault();
            dispatchAction("measure-all");
        }
    });

    var bdSnSetup = document.getElementById("backdropSnSetup");
    [bdInfo, bdLogin, bdPrint, bdAbout, bdApp, bdReportSystemError, bdReportInventory, bdReportHistory].forEach(function (bd) {
        if (!bd) return;
        bd.addEventListener("click", function (e) {
            if (e.target === bd) bd.classList.remove("show");
        });
    });
    bdSnSetup.addEventListener("click", function (e) {
        if (e.target === bdSnSetup) closeSensorNetworkSetup();
    });

    document.getElementById("snSetupSave").addEventListener("click", saveSensorNetworkSetup);
    document.getElementById("snSetupCancel").addEventListener("click", closeSensorNetworkSetup);
    var snSetupCapClose = document.getElementById("snSetupCapClose");
    if (snSetupCapClose) snSetupCapClose.addEventListener("click", closeSensorNetworkSetup);

    var acCapClose = document.getElementById("acCapClose");
    if (acCapClose) acCapClose.addEventListener("click", closeAssignContactsDialog);

    (function wireSimulatedTaskbar() {
        var clock = document.getElementById("taskbarClock");
        if (clock) {
            var timeEl = clock.querySelector(".win-taskbar-time");
            var dateEl = clock.querySelector(".win-taskbar-date");
            function tick() {
                var now = new Date();
                if (timeEl) {
                    timeEl.textContent = now.toLocaleTimeString(undefined, {
                        hour: "numeric",
                        minute: "2-digit"
                    });
                }
                if (dateEl) {
                    dateEl.textContent = now.toLocaleDateString(undefined);
                }
            }
            tick();
            setInterval(tick, 30000);
        }

        var backdrop = document.getElementById("winStartBackdrop");
        var startMenu = document.getElementById("winStartMenu");
        var startBtn = document.getElementById("winStartBtn");
        var winStartSearch = document.getElementById("winStartSearch");
        var winSearchFlyout = document.getElementById("winSearchFlyout");
        var winStartHomeContent = document.getElementById("winStartHomeContent");

        function syncStartSearchView() {
            if (!winSearchFlyout || !winStartHomeContent || !startMenu) return;
            var q = String(winStartSearch && winStartSearch.value ? winStartSearch.value : "").trim().toLowerCase();
            if (q === "services") {
                winSearchFlyout.hidden = false;
                winStartHomeContent.hidden = true;
                startMenu.classList.add("win-start-menu--search");
            } else {
                winSearchFlyout.hidden = true;
                winStartHomeContent.hidden = false;
                startMenu.classList.remove("win-start-menu--search");
            }
        }

        function closeStartMenu() {
            if (!backdrop || !startMenu || !startBtn) return;
            backdrop.hidden = true;
            backdrop.setAttribute("aria-hidden", "true");
            startMenu.hidden = true;
            startBtn.setAttribute("aria-expanded", "false");
            if (winStartSearch) winStartSearch.value = "";
            syncStartSearchView();
            startBtn.focus();
        }

        function openStartMenu() {
            if (!backdrop || !startMenu || !startBtn) return;
            backdrop.hidden = false;
            backdrop.setAttribute("aria-hidden", "false");
            startMenu.hidden = false;
            startBtn.setAttribute("aria-expanded", "true");
            syncStartSearchView();
            if (winStartSearch) {
                setTimeout(function () {
                    winStartSearch.focus();
                }, 0);
            }
        }

        function toggleStartMenu() {
            if (!startMenu) return;
            if (startMenu.hidden) openStartMenu();
            else closeStartMenu();
        }

        if (startBtn) {
            startBtn.addEventListener("click", function (e) {
                e.stopPropagation();
                toggleStartMenu();
            });
        }
        if (backdrop) {
            backdrop.addEventListener("click", closeStartMenu);
        }

        if (winStartSearch) {
            winStartSearch.addEventListener("input", syncStartSearchView);
            winStartSearch.addEventListener("click", function (e) {
                e.stopPropagation();
            });
        }

        if (startMenu) {
            startMenu.addEventListener("click", function (e) {
                var t = e.target;
                var svcBtn = t.closest && t.closest("[data-sim-svc]");
                if (svcBtn) {
                    var act = svcBtn.getAttribute("data-sim-svc");
                    if (act === "admin") {
                        closeStartMenu();
                        if (typeof window.__ebobShowServicesUac === "function") {
                            window.__ebobShowServicesUac();
                        }
                        return;
                    }
                    var svcLabels = {
                        open: "Open",
                        loc: "Open file location",
                        pstart: "Pin to Start",
                        ptask: "Pin to taskbar"
                    };
                    toast((svcLabels[act] || "Action") + " — Services (simulation).");
                    return;
                }
                if (t.closest && t.closest("#winSearchHitServices")) {
                    toast("Services (simulation).");
                    return;
                }
                var tile = t.closest && t.closest(".win-start-tile[data-sim-app]");
                if (tile) {
                    toast(tile.getAttribute("data-sim-app") + " (simulation).");
                    closeStartMenu();
                    return;
                }
                var rec = t.closest && t.closest(".win-start-rec button[data-sim-rec]");
                if (rec) {
                    var k = rec.getAttribute("data-sim-rec");
                    if (k === "ebob") {
                        window.scrollTo({ top: 0, behavior: "smooth" });
                        toast("ebob.html — this troubleshooting page.");
                    } else {
                        var strong = rec.querySelector("strong");
                        toast((strong ? strong.textContent : "Item") + " (simulation).");
                    }
                    closeStartMenu();
                    return;
                }
                var hdrLink = t.closest && t.closest("[data-sim-start]");
                if (hdrLink) {
                    var which = hdrLink.getAttribute("data-sim-start");
                    toast(which === "more" ? "More recommended (simulation)." : "All apps (simulation).");
                    closeStartMenu();
                    return;
                }
            });
        }

        var pins = document.querySelector(".win-taskbar-pins");
        if (!pins) return;
        var btns = pins.querySelectorAll(".win-taskbar-btn");
        if (btns.length < 4) return;
        btns[1].addEventListener("click", function () {
            toast("File Manager (simulation).");
        });
        btns[2].addEventListener("click", function () {
            toast("Browser (simulation).");
        });
        btns[3].addEventListener("click", function () {
            var shell = document.getElementById("appShell");
            if (shell && shell.scrollIntoView) shell.scrollIntoView({ behavior: "smooth", block: "nearest" });
            toast("eBob Workstation — active window.");
        });
    })();

    (function wireServicesMscConsole() {
        var SERVICES_MSC_SEED = [
            { id: "ax", name: "ActiveX Installer (AxInstSV)", description: "Provides User Account Control validation for installation of ActiveX controls.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "appinf", name: "Application Information", description: "Facilitates the running of interactive applications with additional administrative privileges.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "alg", name: "Application Layer Gateway Service", description: "Provides support for third-party protocol plug-ins for Internet Connection Sharing.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "appmgmt", name: "Application Management", description: "Processes installation, removal, and enumeration requests for software deployed through Group Policy.", running: false, startup: "Manual", logOn: "Local System" },
            { id: "audio", name: "Audio", description: "Manages audio for system programs.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "audiosrv", name: "Audio Endpoint Builder", description: "Manages audio devices for the Audio service.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "avctp", name: "AVCTP service", description: "Audio Video Control Transport Protocol service.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "azureattest", name: "AzureAttestService", description: "Supports Azure attestation scenarios.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "bfe", name: "Base Filtering Engine", description: "Base Filtering Engine (BFE) is a service that manages firewall and Internet Protocol security.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "bthserv", name: "Bluetooth Support Service", description: "The Bluetooth service supports discovery and association of remote Bluetooth devices.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "bthusersvc", name: "Bluetooth User Support Service", description: "Bluetooth user session service.", running: false, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "branchcache", name: "BranchCache", description: "This service caches network content from peers on the local subnet.", running: false, startup: "Manual", logOn: "NETWORK SERVICE" },
            { id: "certprop", name: "Certificate Propagation", description: "Copies user certificates and root certificates from smart cards into the user's certificate store.", running: false, startup: "Manual", logOn: "Local System" },
            { id: "clipsvc", name: "Client License Service (ClipSVC)", description: "Enables infrastructure support for the built-in app store.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "complus", name: "COM+ Event System", description: "Supports system event notification for COM+ components.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "diagtrack", name: "Connected User Experiences and Telemetry", description: "The Connected User Experiences and Telemetry service.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "coremessaging", name: "CoreMessaging", description: "Manages communication between system components.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "vaultsvc", name: "Credential Manager", description: "Provides secure storage and retrieval of credentials.", running: true, startup: "Manual", logOn: "Local Service" },
            { id: "cryptsvc", name: "Cryptographic Services", description: "Provides management of certificates, cryptographic keys, and encryption.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "dssvc", name: "Data Sharing Service", description: "Provides data sharing between applications.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "dusmsvc", name: "Data Usage", description: "Network data usage monitoring.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "dcom", name: "DCOM Server Process Launcher", description: "Launches COM and DCOM servers in response to object activation requests.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "das", name: "Device Association Service", description: "Enables pairing between the system and wired or wireless devices.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "dhcp", name: "DHCP Client", description: "Registers and updates IP addresses and DNS records for this computer.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "dps", name: "Diagnostic Policy Service", description: "The Diagnostic Policy Service enables problem detection, troubleshooting, and resolution for system components.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "ebob-engine", name: "eBob Engine Service", description: "Engine service for Binventory 5. Manages sensor communications and measurement processing.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "ebob-scheduler", name: "eBob Scheduler Service", description: "Scheduler service for Binventory 5. Manages schedules and email delivery.", running: true, startup: "Automatic (Delayed Start)", logOn: "Local System" },
            { id: "embedded", name: "Embedded Mode", description: "Enables embedded mode for specialized devices.", running: false, startup: "Manual", logOn: "Local System" },
            { id: "efs", name: "Encrypting File System (EFS)", description: "Provides the core file encryption technology used to store encrypted files on NTFS file system volumes.", running: true, startup: "Manual (Trigger Start)", logOn: "Local System" },
            { id: "entapp", name: "Enterprise App Management Service", description: "Manages enterprise application policies.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "everything", name: "Everything (1.5a)", description: "File search service for the Everything search utility.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "eaphost", name: "Extensible Authentication Protocol", description: "Provides network authentication in such scenarios as 802.1x wired and wireless, VPN, and Network Access Protection (NAP).", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "fhsvc", name: "File History Service", description: "Protects user files from accidental loss by copying them to a backup location.", running: false, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "filesync", name: "FileSyncHelper", description: "Helper service for file synchronization.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "fortemedia", name: "Fortemedia APO Control Service", description: "Audio processing object control.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "fdrespub", name: "Function Discovery Resource Publication", description: "Publishes this computer and its resources so they can be discovered over the network.", running: true, startup: "Manual (Trigger Start)", logOn: "Local Service" },
            { id: "gpsvc", name: "Group Policy Client", description: "Applies Group Policy settings for this computer and users.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "hidserv", name: "Human Interface Device Service", description: "Activates and maintains the use of hot buttons on keyboards, remote controls, and other multimedia devices.", running: false, startup: "Manual", logOn: "Local System" },
            { id: "iphlpsvc", name: "IP Helper", description: "Provides tunnel connectivity using IPv6 transition technologies.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "iphlpsvc2", name: "IKE and AuthIP IPsec Keying Modules", description: "IKEEXT service hosts the Internet Key Exchange and AuthIP keying modules.", running: true, startup: "Automatic (Trigger Start)", logOn: "Local Service" },
            { id: "lanmanserver", name: "Server", description: "Supports file, print, and named-pipe sharing over the network.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "lanmanworkstation", name: "Workstation", description: "Creates and maintains client network connections to remote servers.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "lmhosts", name: "TCP/IP NetBIOS Helper", description: "Provides support for the NetBIOS over TCP/IP (NetBT) service and NetBIOS name resolution for clients on the network.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "mpssvc", name: "Defender Firewall", description: "Helps protect your PC by blocking unauthorized access.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "netman", name: "Network Connections", description: "Manages objects in the Network and Dial-Up Connections folder.", running: true, startup: "Manual", logOn: "Local System" },
            { id: "netprofm", name: "Network List Service", description: "Identifies the networks to which the computer has connected.", running: true, startup: "Manual", logOn: "Local Service" },
            { id: "nlasvc", name: "Network Location Awareness", description: "Collects and stores configuration information for the network.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "nsi", name: "Network Store Interface Service", description: "This service delivers network notifications.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "pla", name: "Performance Logs & Alerts", description: "Collects performance data from local or remote computers.", running: false, startup: "Manual", logOn: "Local Service" },
            { id: "plugplay", name: "Plug and Play", description: "Enables a computer to recognize and adapt to hardware changes with minimal user intervention.", running: true, startup: "Manual", logOn: "Local System" },
            { id: "power", name: "Power", description: "Manages power policy and power policy notification delivery.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "rasauto", name: "Remote Access Auto Connection Manager", description: "Creates a connection to a remote network whenever a program references a remote DNS or NetBIOS name.", running: false, startup: "Manual", logOn: "Local System" },
            { id: "rpcss", name: "Remote Procedure Call (RPC)", description: "The RPCSS service is the Service Control Manager for COM and DCOM servers.", running: true, startup: "Automatic", logOn: "Network Service" },
            { id: "schedule", name: "Task Scheduler", description: "Enables a user to configure and schedule automated tasks on this computer.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "spooler", name: "Print Spooler", description: "This service spools print jobs and handles interaction with the printer.", running: true, startup: "Automatic", logOn: "Local System" },
            { id: "w32time", name: "Time", description: "Maintains date and time synchronization on all clients and servers in the network.", running: true, startup: "Automatic", logOn: "Local Service" },
            { id: "wuauserv", name: "Update", description: "Enables the detection, download, and installation of updates for the system and other programs.", running: true, startup: "Manual (Trigger Start)", logOn: "Local System" },
            { id: "wsearch", name: "Search", description: "Provides content indexing, property caching, and search results for files, e-mail, and other content.", running: true, startup: "Automatic (Delayed Start)", logOn: "Local System" }
        ];

        var backdropUac = document.getElementById("backdropUac");
        var backdropSvc = document.getElementById("backdropServicesMsc");
        var tbody = document.getElementById("svcMscTbody");
        var uacYes = document.getElementById("uacYes");
        var uacNo = document.getElementById("uacNo");
        var svcMscClose = document.getElementById("svcMscClose");
        var svcMscMin = document.getElementById("svcMscMin");
        var svcMscMax = document.getElementById("svcMscMax");
        var svcTbStart = document.getElementById("svcTbStart");
        var svcTbStop = document.getElementById("svcTbStop");
        var svcTbRestart = document.getElementById("svcTbRestart");
        var svcTbRefresh = document.getElementById("svcTbRefresh");
        var svcTbPause = document.getElementById("svcTbPause");
        var svcTabExtended = document.getElementById("svcTabExtended");
        var svcTabStandard = document.getElementById("svcTabStandard");
        var svcMscContentInner = document.getElementById("svcMscContentInner");
        var svcMscDescTitle = document.getElementById("svcMscDescTitle");
        var svcMscDescBody = document.getElementById("svcMscDescBody");
        var svcMscDescLinks = document.getElementById("svcMscDescLinks");
        var svcCtxMenu = document.getElementById("svcCtxMenu");
        var svcCtxStart = document.getElementById("svcCtxStart");
        var svcCtxStop = document.getElementById("svcCtxStop");
        var svcCtxRestart = document.getElementById("svcCtxRestart");

        var mscState = {
            services: [],
            selectedId: null,
            eBobCycle: 0,
            ctxServiceId: null
        };

        var EBOB_IDS = ["ebob-engine", "ebob-scheduler"];

        function cloneSeed() {
            return SERVICES_MSC_SEED.map(function (s) {
                return {
                    id: s.id,
                    name: s.name,
                    description: s.description,
                    running: s.running,
                    startup: s.startup,
                    logOn: s.logOn
                };
            });
        }

        function getService(id) {
            for (var i = 0; i < mscState.services.length; i++) {
                if (mscState.services[i].id === id) return mscState.services[i];
            }
            return null;
        }

        function hideCtxMenu() {
            if (svcCtxMenu) {
                svcCtxMenu.hidden = true;
                svcCtxMenu.setAttribute("aria-hidden", "true");
            }
            mscState.ctxServiceId = null;
        }

        function showUac() {
            if (!backdropUac) return;
            backdropUac.classList.add("show");
            backdropUac.setAttribute("aria-hidden", "false");
            setTimeout(function () {
                if (uacNo) uacNo.focus();
            }, 0);
        }

        function hideUac() {
            if (!backdropUac) return;
            backdropUac.removeAttribute("data-uac-context");
            backdropUac.classList.remove("show");
            backdropUac.setAttribute("aria-hidden", "true");
        }

        function syncMscEngineToWorkstation() {
            var eng = getService("ebob-engine");
            if (!eng) return;
            setEbobServicesRunningFromSim(eng.running);
        }

        function showServicesMsc() {
            if (!backdropSvc) return;
            mscState.services = cloneSeed();
            var engSync = getService("ebob-engine");
            var schSync = getService("ebob-scheduler");
            if (state && state.ebobServicesRunning === false) {
                if (engSync) engSync.running = false;
                if (schSync) schSync.running = false;
            }
            mscState.selectedId = mscState.services.length ? mscState.services[0].id : null;
            mscState.eBobCycle = 0;
            renderServicesTable();
            updateDescAndToolbar();
            backdropSvc.classList.add("show");
            backdropSvc.setAttribute("aria-hidden", "false");
            hideCtxMenu();
            setTimeout(function () {
                var wrap = document.querySelector(".svc-msc-table-wrap");
                if (wrap) wrap.focus();
            }, 0);
        }

        function hideServicesMsc() {
            if (!backdropSvc) return;
            backdropSvc.classList.remove("show");
            backdropSvc.setAttribute("aria-hidden", "true");
            hideCtxMenu();
            syncMscEngineToWorkstation();
        }

        function statusText(s) {
            return s.running ? "Running" : "";
        }

        function renderServicesTable() {
            if (!tbody) return;
            tbody.innerHTML = "";
            for (var i = 0; i < mscState.services.length; i++) {
                var s = mscState.services[i];
                var tr = document.createElement("tr");
                tr.className = "svc-row" + (mscState.selectedId === s.id ? " svc-row-selected" : "");
                tr.setAttribute("data-service-id", s.id);
                tr.setAttribute("role", "row");
                var tdName = document.createElement("td");
                tdName.textContent = s.name;
                var tdDesc = document.createElement("td");
                tdDesc.textContent = s.description;
                var tdStat = document.createElement("td");
                tdStat.textContent = statusText(s);
                tdStat.className = "svc-col-status";
                var tdStart = document.createElement("td");
                tdStart.textContent = s.startup;
                var tdLog = document.createElement("td");
                tdLog.textContent = s.logOn;
                tr.appendChild(tdName);
                tr.appendChild(tdDesc);
                tr.appendChild(tdStat);
                tr.appendChild(tdStart);
                tr.appendChild(tdLog);
                tbody.appendChild(tr);
            }
        }

        function syncRowSelectionClass() {
            var rows = tbody ? tbody.querySelectorAll(".svc-row") : [];
            for (var r = 0; r < rows.length; r++) {
                var row = rows[r];
                var id = row.getAttribute("data-service-id");
                if (id === mscState.selectedId) row.classList.add("svc-row-selected");
                else row.classList.remove("svc-row-selected");
            }
        }

        function selectService(id) {
            if (!getService(id)) return;
            mscState.selectedId = id;
            syncRowSelectionClass();
            updateDescAndToolbar();
        }

        function scrollRowIntoView(id) {
            if (!tbody) return;
            var row = tbody.querySelector('[data-service-id="' + id + '"]');
            if (row && row.scrollIntoView) row.scrollIntoView({ block: "nearest", behavior: "smooth" });
        }

        function updateDescAndToolbar() {
            var s = mscState.selectedId ? getService(mscState.selectedId) : null;
            if (svcMscDescTitle) svcMscDescTitle.textContent = s ? s.name : "—";
            if (svcMscDescBody) svcMscDescBody.textContent = s ? s.description : "Select a service in the list to see its description.";
            if (svcMscDescLinks) svcMscDescLinks.hidden = !s;

            var canStart = s && !s.running;
            var canStop = s && s.running;
            if (svcTbStart) svcTbStart.disabled = !canStart;
            if (svcTbStop) svcTbStop.disabled = !canStop;
            if (svcTbRestart) svcTbRestart.disabled = !s;
            if (svcTbPause) svcTbPause.disabled = true;
        }

        function applyStatusToRow(id) {
            var s = getService(id);
            if (!s) return;
            var row = tbody ? tbody.querySelector('[data-service-id="' + id + '"]') : null;
            if (row) {
                var cells = row.querySelectorAll("td");
                if (cells.length > 2) cells[2].textContent = statusText(s);
            }
        }

        /** eBob: starting Scheduler also starts Engine; stopping Engine also stops Scheduler; starting Engine also starts Scheduler. */
        function startService(id) {
            var s = getService(id);
            if (!s || s.running) return;
            s.running = true;
            applyStatusToRow(id);
            if (id === "ebob-engine") {
                var sch = getService("ebob-scheduler");
                if (sch && !sch.running) {
                    sch.running = true;
                    applyStatusToRow("ebob-scheduler");
                }
            }
            if (id === "ebob-scheduler") {
                var eng = getService("ebob-engine");
                if (eng && !eng.running) {
                    eng.running = true;
                    applyStatusToRow("ebob-engine");
                }
            }
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
        }

        function stopService(id) {
            var s = getService(id);
            if (!s || !s.running) return;
            s.running = false;
            applyStatusToRow(id);
            if (id === "ebob-engine") {
                var sch = getService("ebob-scheduler");
                if (sch && sch.running) {
                    sch.running = false;
                    applyStatusToRow("ebob-scheduler");
                }
            }
            if (id === "ebob-scheduler") {
                var eng = getService("ebob-engine");
                if (eng && eng.running) {
                    eng.running = false;
                    applyStatusToRow("ebob-engine");
                }
            }
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
        }

        function restartService(id) {
            var s = getService(id);
            if (!s) return;
            if (s.running) {
                s.running = false;
                applyStatusToRow(id);
                if (id === "ebob-engine") {
                    var sch = getService("ebob-scheduler");
                    if (sch && sch.running) {
                        sch.running = false;
                        applyStatusToRow("ebob-scheduler");
                    }
                }
                if (id === "ebob-scheduler") {
                    var engR = getService("ebob-engine");
                    if (engR && engR.running) {
                        engR.running = false;
                        applyStatusToRow("ebob-engine");
                    }
                }
                updateDescAndToolbar();
                syncMscEngineToWorkstation();
                setTimeout(function () {
                    startService(id);
                }, 250);
            } else {
                startService(id);
            }
        }

        function positionCtxMenu(clientX, clientY) {
            if (!svcCtxMenu) return;
            svcCtxMenu.style.left = Math.min(clientX, window.innerWidth - 180) + "px";
            svcCtxMenu.style.top = Math.min(clientY, window.innerHeight - 200) + "px";
        }

        function updateCtxMenuState() {
            var s = mscState.ctxServiceId ? getService(mscState.ctxServiceId) : null;
            if (svcCtxStart) svcCtxStart.disabled = !s || s.running;
            if (svcCtxStop) svcCtxStop.disabled = !s || !s.running;
            if (svcCtxRestart) svcCtxRestart.disabled = !s;
        }

        function onServicesKeydown(e) {
            if (!backdropSvc || !backdropSvc.classList.contains("show")) return;
            var tag = e.target && e.target.tagName;
            if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
            if (e.key === "Escape") {
                e.preventDefault();
                e.stopPropagation();
                hideServicesMsc();
                return;
            }
            if (e.key === "e" || e.key === "E") {
                e.preventDefault();
                var id = EBOB_IDS[mscState.eBobCycle % EBOB_IDS.length];
                mscState.eBobCycle += 1;
                selectService(id);
                scrollRowIntoView(id);
            }
        }

        function onGlobalEscape(e) {
            if (e.key !== "Escape") return;
            if (backdropUac && backdropUac.classList.contains("show")) {
                e.preventDefault();
                hideUac();
                return;
            }
        }

        if (tbody) {
            tbody.addEventListener("click", function (e) {
                var tr = e.target.closest && e.target.closest("tr[data-service-id]");
                if (!tr) return;
                selectService(tr.getAttribute("data-service-id"));
            });
            tbody.addEventListener("contextmenu", function (e) {
                var tr = e.target.closest && e.target.closest("tr[data-service-id]");
                if (!tr) return;
                e.preventDefault();
                var id = tr.getAttribute("data-service-id");
                selectService(id);
                mscState.ctxServiceId = id;
                if (svcCtxMenu) {
                    updateCtxMenuState();
                    svcCtxMenu.hidden = false;
                    svcCtxMenu.setAttribute("aria-hidden", "false");
                    positionCtxMenu(e.clientX, e.clientY);
                }
            });
        }

        document.addEventListener("keydown", onServicesKeydown, true);
        document.addEventListener("keydown", onGlobalEscape, true);

        document.addEventListener("click", function (e) {
            if (!svcCtxMenu || svcCtxMenu.hidden) return;
            if (e.target.closest && e.target.closest("#svcCtxMenu")) return;
            hideCtxMenu();
        }, true);

        if (uacYes) {
            uacYes.addEventListener("click", function () {
                var ctx = backdropUac ? backdropUac.getAttribute("data-uac-context") : null;
                hideUac();
                if (ctx === "import-restart") {
                    completeSimImportRestart();
                    return;
                }
                showServicesMsc();
            });
        }
        if (uacNo) {
            uacNo.addEventListener("click", function () {
                if (backdropUac) backdropUac.removeAttribute("data-uac-context");
                hideUac();
            });
        }

        if (backdropUac) {
            backdropUac.addEventListener("click", function (e) {
                if (e.target === backdropUac) hideUac();
            });
        }

        if (svcMscClose) svcMscClose.addEventListener("click", hideServicesMsc);
        if (svcMscMin) svcMscMin.addEventListener("click", hideServicesMsc);
        if (svcMscMax) svcMscMax.addEventListener("click", function () {
            toast("Maximize (simulation).");
        });

        if (svcTbStart) svcTbStart.addEventListener("click", function () {
            if (mscState.selectedId) startService(mscState.selectedId);
        });
        if (svcTbStop) svcTbStop.addEventListener("click", function () {
            if (mscState.selectedId) stopService(mscState.selectedId);
        });
        if (svcTbRestart) svcTbRestart.addEventListener("click", function () {
            if (mscState.selectedId) restartService(mscState.selectedId);
        });
        if (svcTbRefresh) svcTbRefresh.addEventListener("click", function () {
            mscState.services = cloneSeed();
            var engSync = getService("ebob-engine");
            var schSync = getService("ebob-scheduler");
            if (state && state.ebobServicesRunning === false) {
                if (engSync) engSync.running = false;
                if (schSync) schSync.running = false;
            }
            renderServicesTable();
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
            toast("Refreshed (simulation).");
        });

        if (svcTabExtended && svcTabStandard && svcMscContentInner) {
            svcTabExtended.addEventListener("click", function () {
                svcMscContentInner.classList.remove("svc-msc-standard");
                svcTabExtended.classList.add("svc-tab-active");
                svcTabStandard.classList.remove("svc-tab-active");
                svcTabExtended.setAttribute("aria-selected", "true");
                svcTabStandard.setAttribute("aria-selected", "false");
            });
            svcTabStandard.addEventListener("click", function () {
                svcMscContentInner.classList.add("svc-msc-standard");
                svcTabStandard.classList.add("svc-tab-active");
                svcTabExtended.classList.remove("svc-tab-active");
                svcTabStandard.setAttribute("aria-selected", "true");
                svcTabExtended.setAttribute("aria-selected", "false");
            });
        }

        if (svcMscDescLinks) {
            svcMscDescLinks.addEventListener("click", function (e) {
                var a = e.target.closest && e.target.closest("[data-svc-link]");
                if (!a || !mscState.selectedId) return;
                e.preventDefault();
                var act = a.getAttribute("data-svc-link");
                if (act === "start") startService(mscState.selectedId);
                else if (act === "stop") stopService(mscState.selectedId);
                else if (act === "restart") restartService(mscState.selectedId);
            });
        }

        function ctxAction(fn) {
            var id = mscState.ctxServiceId || mscState.selectedId;
            if (id) fn(id);
            hideCtxMenu();
        }

        if (svcCtxStart) svcCtxStart.addEventListener("click", function () { ctxAction(startService); });
        if (svcCtxStop) svcCtxStop.addEventListener("click", function () { ctxAction(stopService); });
        if (svcCtxRestart) svcCtxRestart.addEventListener("click", function () { ctxAction(restartService); });

        if (backdropSvc) {
            backdropSvc.addEventListener("click", function (e) {
                if (e.target === backdropSvc) hideServicesMsc();
            });
        }

        window.__ebobShowServicesUac = showUac;
    })();
})();
