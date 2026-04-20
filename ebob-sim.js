(function () {
    "use strict";

    /** Legacy key; storage is cleared on load and not written (session-only state). */
    var STORAGE_KEY = "ebobSimState.v1";
    var GRID_SIZE = 16;
    /** Invalidates in-flight staggered vessel reveals when renderGrid runs again. */
    var vesselGridRevealGen = 0;
    /** BLL frmVesselSetup — txtVesselName.MaxLength = 10; Contents max 25 (BLL.vb). */
    var EBOB_VESSEL_NAME_MAX = 10;
    var EBOB_VESSEL_CONTENTS_MAX = 25;

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

    /** Persists eBob Engine / Scheduler running flags across reload and Binventory relaunch from desktop (sessionStorage). */
    var EBOB_SVC_SESSION_KEY = "ebobTutorialSvcStateV1";
    var EBOB_TUTORIAL_MODE_KEY = "ebobTutorialModeV1";
    var EBOB_TUTORIAL_STATE_KEY = "ebobTutorialScenarioStateV1";
    var EBOB_TUTORIAL_MODE_DB_READ_ONLY = "db-read-only";
    var EBOB_TUTORIAL_MODE_PENDING_UNKNOWN = "pending-unknown-status";
    var EBOB_TUTORIAL_MODE_FREE = "free-mode";
    var SIM_WORKSTATION_IPV4 = "10.101.70.69";
    var activeTutorialMode = null;
    var dbReadOnlyTutorial = null;
    var pendingUnknownTutorial = null;
    var dbTutorialGuardsBound = false;

    function persistEbobSvcSession() {
        try {
            sessionStorage.setItem(
                EBOB_SVC_SESSION_KEY,
                JSON.stringify({
                    engineRunning: !!state.ebobServicesRunning,
                    schedulerRunning: !!state.ebobSchedulerRunning
                })
            );
        } catch (e) {
            /* private mode */
        }
    }

    /**
     * After relaunch from desktop (same tab, no full reload): restore engine/scheduler from session.
     * Scheduler cannot run without engine — clamped here. Full page refresh clears session (loadState).
     */
    function applyEbobSvcSession() {
        try {
            var raw = sessionStorage.getItem(EBOB_SVC_SESSION_KEY);
            if (!raw) return;
            var s = JSON.parse(raw);
            if (typeof s.engineRunning !== "boolean") return;
            state.ebobServicesRunning = s.engineRunning;
            state.ebobSchedulerRunning =
                typeof s.schedulerRunning === "boolean" ? s.schedulerRunning : s.engineRunning;
            if (!state.ebobServicesRunning) state.ebobSchedulerRunning = false;
        } catch (e2) {
            /* ignore */
        }
    }

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
        return map[unitName || ""] || "cu ft";
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
        return map[unitName || ""] || "lbs";
    }

    /** Convert internal cubic feet to the vessel's display volume unit (same basis as eBob inventory math). */
    function volumeFromCuFt(cuFt, unitName) {
        var v = Number(cuFt);
        if (isNaN(v)) v = 0;
        var u = unitName || "Cubic Feet";
        if (u === "U.S. Gallon (Liquid)") return v * 7.48051948;
        if (u === "Cubic Feet") return v;
        if (u === "Cubic Yard") return v / 27;
        if (u === "Imperial Gallon") return v * 6.22883545;
        if (u === "Liter") return v * 28.316846592;
        if (u === "U.S. Oil Barrel (42 gal)") return (v * 7.48051948) / 42;
        if (u === "Cubic Meter") return v * 0.028316846592;
        if (u === "U.S. Dry Gallon") return v / 0.155556502;
        return v;
    }

    /** Convert internal pounds to the vessel's display weight unit. */
    function weightFromLb(lb, unitName) {
        var w = Number(lb);
        if (isNaN(w)) w = 0;
        var u = unitName || "Pounds";
        if (u === "Tons") return w / 2000;
        if (u === "Pounds") return w;
        if (u === "Kilograms") return w * 0.45359237;
        if (u === "Metric Tonnes") return w / 2204.62262185;
        if (u === "Ounces") return w * 16;
        if (u === "Hundredweight (US)") return w / 100;
        return w;
    }

    function formatVolumeWeightPair(v, cuFtVal, lbVal) {
        ensureVesselSetupDefaults(v);
        var volU = volumeFromCuFt(cuFtVal, v.volumeDisplayUnits);
        var wtU = weightFromLb(lbVal, v.weightDisplayUnits);
        var volAbbr = volumeAbbrevForDisplay(v.volumeDisplayUnits);
        var wtAbbr = weightAbbrevForDisplay(v.weightDisplayUnits);
        return {
            volStr: volU.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " " + volAbbr,
            wtStr: wtU.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " " + wtAbbr
        };
    }

    function parseAlarmThresholdPct(raw) {
        if (raw == null) return null;
        var s = String(raw).trim().replace(/%/g, "");
        if (s === "") return null;
        var n = parseFloat(s);
        if (isNaN(n)) return null;
        return Math.max(0, Math.min(100, n));
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
        var tb = vesselSetupField("vs_custom_tbody");
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
        var sel = vesselSetupField("vs_custom_output_type");
        if (sel) {
            v.customOutputTypeIndex = parseInt(sel.value, 10);
            if (isNaN(v.customOutputTypeIndex)) v.customOutputTypeIndex = 0;
        }
    }

    function sortCustomStrapTbodyByDistance() {
        var tb = vesselSetupField("vs_custom_tbody");
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
        var tb3 = vesselSetupField("vs_custom_tbody");
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
                : vesselSetupField("vs_vessel_type")
                  ? parseInt(vesselSetupField("vs_vessel_type").value, 10) || 1
                  : parseInt(v.vesselTypeId, 10) || 1;
        if (tid === 15) {
            refreshVesselCustomStrapFromDom(v);
            return;
        }
        ensureVesselShapeParams(v);
        var i;
        for (i = 0; i < 7; i++) {
            var el = vesselSetupField("vs_sp_" + i);
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
            /* tblMeasurementData_Temp — per tblDevice address (frmVesselDetails RefreshSmartBobA) */
            if (sid === 1 || sid === 3) {
                ensureVesselSetupDefaults(v);
                v.sbMeasurementTempByAddr = v.sbMeasurementTempByAddr || {};
                var m = vesselDisplayMetrics(v);
                var abbr =
                    state.systemSettings && state.systemSettings.units === "Metric" ? "m" : "ft";
                smartBobADeviceRows(v).forEach(function (row) {
                    if (!row.isActive) return;
                    v.sbMeasurementTempByAddr[String(row.address)] = {
                        headroomHeight: m.headroom.heightFt,
                        productHeight: m.product.heightFt,
                        distanceAbbrev: abbr
                    };
                });
            }
            if (sid === 2) {
                v.smartBobDropCount = (parseInt(v.smartBobDropCount, 10) || 0) + 1;
                v.smartBobRetractCount = (parseInt(v.smartBobRetractCount, 10) || 0) + 1;
            }
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

    /** Each name must be unique after clamp to EBOB_VESSEL_NAME_MAX (10); write short labels here, do not rely on truncation. */
    var VESSEL_NAMES = [
        "Cem Silo 1", "Cem Silo 2", "Fly Ash E", "Fly Ash W",
        "Sand Bin A", "Sand Bin B", "Agg Bin 3", "Agg Bin 4",
        "Chem Tk 1", "Chem Tk 2", "Flour Silo", "Sugar Silo",
        "Pellet Day", "Pellet Nite", "Reserve 1", "Reserve 2",
        "Overflow A", "Overflow B", "Staging N", "Staging S",
        "Mill Feed", "Coarse Bin", "Fine Bin", "Dust Clctr",
        "Silo 25", "Silo 26", "Silo 27", "Silo 28",
        "Tank Cold", "Tank Hot", "Mix Upper", "Mix Lower"
    ];

    var state = {
        currentUser: null,
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
            /* frmSystem — "Do NOT require users to log in" (AutoLoginFlag); on = skip login / implicit admin session. */
            autoLogin: true
        },
        users: [
            {
                id: "u1",
                userId: "admin",
                name: "System Administrator",
                role: "Administrator",
                accessLevel: 1,
                authenticationMethod: 0,
                password: "1234",
                jobTitle: "System Administrator",
                userType: "Administrator User",
                firstName: "System",
                middleInit: "",
                lastName: "Administrator",
                lastLogon: ""
            }
        ],
        emailReportPrefs: { to: "operator@example.com", frequency: "Daily" },
        currentWorkstationSiteId: "st1",
        /** eBob Engine running — BobMsgQue reachable (Vessel.vb GetSensorStatus). */
        ebobServicesRunning: true,
        /** Scheduler service running (services.msc); independent of engine for start/stop. */
        ebobSchedulerRunning: true,
        /**
         * Mirrors AppGlobals.vb gbVesselsReadOnly — latched when engine stops in-session; on reload, cleared only if engine is up.
         */
        vesselsReadOnly: false
    };

    function findCurrentSite() {
        var sid = state.currentWorkstationSiteId;
        if (!sid || !state.sites || !state.sites.length) return null;
        for (var i = 0; i < state.sites.length; i++) {
            if (state.sites[i].id === sid) return state.sites[i];
        }
        return null;
    }

    function isCurrentSiteHostMismatch() {
        var s = findCurrentSite();
        if (!s) return false;
        ensureSiteFields(s);
        var host = String(s.serviceHostIp || "").trim().toLowerCase();
        if (!host) return false;
        if (host === "127.0.0.1" || host === "localhost") return false;
        return host !== SIM_WORKSTATION_IPV4;
    }

    function syncReadOnlyLatchOnStartup() {
        state.vesselsReadOnly = !state.ebobServicesRunning || isCurrentSiteHostMismatch();
    }

    function getReadOnlyMessage() {
        if (!state.ebobServicesRunning) {
            return "Database is read only — eBob Engine is not running.";
        }
        if (isCurrentSiteHostMismatch()) {
            return (
                "Database is read only — Site Maintenance Host IP does not match this workstation IPv4 (" +
                SIM_WORKSTATION_IPV4 +
                "). Update Host IP and restart Binventory."
            );
        }
        return "Database is read only — close and restart Binventory to restore full access.";
    }

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
        var prod = PRODUCTS[i % PRODUCTS.length];
        /** Same fields as Vessel Setup → Alarms (not decorative): high/low enabled + % Full thresholds. */
        var alarmHighEnabled = false;
        var alarmHighPct = "";
        var alarmLowEnabled = false;
        var alarmLowPct = "";
        var mod6 = i % 6;
        if (mod6 === 0 || mod6 === 2 || mod6 === 4) {
            alarmHighEnabled = true;
            alarmHighPct = "85";
            alarmLowEnabled = true;
            alarmLowPct = "15";
            if (mod6 === 0) {
                pct = 91;
            } else if (mod6 === 2) {
                pct = 9;
            }
        }
        var vol = Math.round(1000 + i * 33);
        if (alarmHighEnabled && (mod6 === 0 || mod6 === 2)) {
            vol = Math.round(800 + pct * 45);
        }
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
            fillColor: "Dark Red",
            volumeDisplayUnits: "Cubic Feet",
            weightDisplayUnits: "Pounds",
            defaultHeadroom: false,
            vesselTypeId: 1,
            sensorTypeId: 1,
            densityMode: "density",
            productDensity: "40",
            densityUnits: "lbs / cubic ft",
            specificGravity: "",
            alarmPreLowEnabled: false,
            alarmPreLowPct: "",
            alarmLowEnabled: alarmLowEnabled,
            alarmLowPct: alarmLowPct,
            alarmHighEnabled: alarmHighEnabled,
            alarmHighPct: alarmHighPct,
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
            {
                id: "c1",
                firstName: "Jane",
                lastName: "Operator",
                name: "Jane Operator",
                email: "jane@example.com",
                emailAddress: "jane@example.com",
                phone: "555-0100",
                jobTitle: ""
            },
            {
                id: "c2",
                firstName: "John",
                lastName: "Smith",
                name: "John Smith",
                email: "john@example.com",
                emailAddress: "john@example.com",
                phone: "555-0102",
                jobTitle: ""
            }
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
        state.currentUser = null;
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
                userId: "admin",
                name: "System Administrator",
                role: "Administrator",
                accessLevel: 1,
                authenticationMethod: 0,
                password: "1234",
                jobTitle: "System Administrator",
                userType: "Administrator User",
                firstName: "System",
                middleInit: "",
                lastName: "Administrator",
                lastLogon: ""
            }
        ];
        state.emailReportPrefs = { to: "operator@example.com", frequency: "Daily" };
        state.currentWorkstationSiteId = "st1";
        state.ebobServicesRunning = true;
        state.ebobSchedulerRunning = true;
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
        clampVesselStringsFromEbobLimits(v);
        if (v.status === "Idle" || v.status === "OK") {
            v.status = defaultIdleStatusForSensorType(v.sensorTypeId);
        }
        if (!vesselHasValidSavedSensorAddress(v)) {
            assignUniqueSensorAddressForVessel(v);
        }
    }

    function ensureVesselSetupDefaults(v) {
        if (v.fillColor == null) v.fillColor = "Dark Red";
        if (v.volumeDisplayUnits == null) v.volumeDisplayUnits = "Cubic Feet";
        if (v.weightDisplayUnits == null) v.weightDisplayUnits = "Pounds";
        if (v.defaultHeadroom == null) v.defaultHeadroom = !!v.headroom;
        if (v.headroom == null) v.headroom = !!v.defaultHeadroom;
        if (v.capacityHeightFt == null) v.capacityHeightFt = 14;
        if (v.vesselTypeId == null) v.vesselTypeId = 1;
        if (v.sensorTypeId == null) v.sensorTypeId = 2;
        if (v.densityMode == null) v.densityMode = "density";
        if (v.productDensity == null) v.productDensity = "40";
        if (v.densityUnits == null) v.densityUnits = "lbs / cubic ft";
        if (v.specificGravity == null) v.specificGravity = "";
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
        /* tblDevice.Label / RelativeWeight (0–1) — frmVesselDetails SmartBob A */
        if (v.deviceLabel == null) v.deviceLabel = "";
        if (v.deviceRelativeWeight == null) v.deviceRelativeWeight = "1";
        /* tblMeasurementData_Temp equivalent: per-address headroom/product after measure */
        if (v.sbMeasurementTempByAddr == null || typeof v.sbMeasurementTempByAddr !== "object") {
            v.sbMeasurementTempByAddr = {};
        }
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

    /**
     * After load or seed, two vessels may still share a valid address (e.g. both "1").
     * Reassign in stable order until each address on a network is unique — mirrors BLL duplicate checks.
     */
    function resolveDuplicateSensorAddressesOnAllVessels() {
        var guard = 0;
        var changed = true;
        while (changed && guard < 64) {
            guard += 1;
            changed = false;
            var usage = {};
            state.vessels.forEach(function (v) {
                ensureVesselSetupDefaults(v);
                var nid = v.sensorNetworkId;
                if (!nid) return;
                if (!usage[nid]) usage[nid] = {};
                var u = usage[nid];
                var addrs = vesselDeviceAddressStrings(v);
                var conflict = false;
                var ai;
                for (ai = 0; ai < addrs.length; ai++) {
                    var a = addrs[ai];
                    if (u[a] && u[a] !== v.id) conflict = true;
                }
                if (conflict) {
                    assignUniqueSensorAddressForVessel(v);
                    changed = true;
                    addrs = vesselDeviceAddressStrings(v);
                }
                for (ai = 0; ai < addrs.length; ai++) {
                    u[addrs[ai]] = v.id;
                }
            });
        }
    }

    /** Clamp persisted strings to BLL limits (defensive for imports / old seeds). */
    function clampVesselStringsFromEbobLimits(v) {
        if (!v) return;
        if (v.name != null && String(v.name).length > EBOB_VESSEL_NAME_MAX) {
            v.name = String(v.name).slice(0, EBOB_VESSEL_NAME_MAX);
        }
        if (v.contents != null && String(v.contents).length > EBOB_VESSEL_CONTENTS_MAX) {
            v.contents = String(v.contents).slice(0, EBOB_VESSEL_CONTENTS_MAX);
        }
        if (v.product != null && String(v.product).length > EBOB_VESSEL_CONTENTS_MAX) {
            v.product = String(v.product).slice(0, EBOB_VESSEL_CONTENTS_MAX);
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

    /**
     * Vessel Setup form: either stacked (#appModalBodyStack) over Vessel Maintenance, or legacy base (#appModalBody).
     */
    function vesselSetupDomRoot() {
        var bdStack = document.getElementById("backdropAppStack");
        var stackBody = document.getElementById("appModalBodyStack");
        if (
            bdStack &&
            bdStack.classList.contains("show") &&
            stackBody &&
            (stackBody.querySelector(".vs-shell") || stackBody.querySelector("#vs_name"))
        ) {
            return stackBody;
        }
        return document.getElementById("appModalBody");
    }

    function vesselSetupField(id) {
        var root = vesselSetupDomRoot();
        if (root && root.querySelector) {
            var el = root.querySelector("#" + id);
            if (el) return el;
        }
        return document.getElementById(id);
    }

    /** Read frmVesselSetup Contents tab — cboContents is editable ComboBox (not DropDownList). */
    function getVsContentsTextFromDom() {
        var el = vesselSetupField("vs_contents");
        if (!el) return "";
        return String(el.value != null ? el.value : "").trim();
    }

    /** BLL ValidateVesselSetup — display tab before sensor block. */
    function validateVesselSetupDisplayFromDom() {
        if (!vmSetupEditingId) {
            return {
                ok: false,
                message: "Vessel Setup session is not active. Close the dialog and open Vessel Setup again."
            };
        }
        var nmEl = vesselSetupField("vs_name");
        var nameVal = nmEl ? String(nmEl.value).trim() : "";
        if (nameVal.length === 0 || nameVal.length > EBOB_VESSEL_NAME_MAX) {
            return {
                ok: false,
                message: "Vessel Name is required and must be specified in 10 characters or less."
            };
        }
        var contentsVal = getVsContentsTextFromDom();
        if (contentsVal.length === 0 || contentsVal.length > EBOB_VESSEL_CONTENTS_MAX) {
            return {
                ok: false,
                message: "Contents is required and must be specified in 25 characters or less."
            };
        }
        var vv = findVessel(vmSetupEditingId);
        var siteId = vv && vv.siteId != null ? vv.siteId : state.currentWorkstationSiteId;
        var other = findOtherVesselWithNameAtSite(vmSetupEditingId, siteId, nameVal);
        if (other) {
            return {
                ok: false,
                message: "Vessel Name is already in use at this site. Each vessel must have a unique name."
            };
        }
        return { ok: true };
    }

    function findOtherVesselWithNameAtSite(excludeVesselId, siteId, nameTrimmed) {
        var nt = String(nameTrimmed || "").toLowerCase();
        var siteNorm =
            siteId != null && siteId !== "" ? String(siteId) : String(state.currentWorkstationSiteId || "");
        var ex = excludeVesselId != null ? String(excludeVesselId) : "";
        var vi;
        for (vi = 0; vi < state.vessels.length; vi++) {
            var ov = state.vessels[vi];
            if (ex !== "" && String(ov.id) === ex) continue;
            var vst =
                ov.siteId != null && ov.siteId !== ""
                    ? String(ov.siteId)
                    : String(state.currentWorkstationSiteId || "");
            if (vst !== siteNorm) continue;
            var on = ov.name != null ? String(ov.name).trim().toLowerCase() : "";
            if (on === nt) return ov;
        }
        return null;
    }

    function validateVesselSetupSensorBlockFromDom() {
        var vid = vmSetupEditingId;
        if (!vid) {
            return {
                ok: false,
                message: "Vessel Setup session is not active. Close the dialog and open Vessel Setup again."
            };
        }
        var netEl = vesselSetupField("vs_network");
        var stEl = vesselSetupField("vs_sensor_type");
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
            var vcEl = vesselSetupField("vs_vega_count");
            var n = Math.min(32, Math.max(1, parseInt(vcEl && vcEl.value, 10) || 1));
            var seen = {};
            var i;
            for (i = 1; i <= n; i++) {
                var aEl = vesselSetupField("vs_vg_a" + i);
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
            var sbAddr = vesselSetupField("vs_sb_addr1");
            var raw = sbAddr ? sbAddr.value : "";
            var vr = validateIntegerSensorAddress(raw, b.min, b.max);
            if (!vr.ok) return { ok: false, message: vr.msg };
            pending.push(vr.normalized);
        } else {
            var genAddr = vesselSetupField("vs_sensor_address");
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
            if (excludeVesselId != null && ov.id === excludeVesselId) continue;
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

    /** Capacity volume/weight from current product fill (same math as vesselDisplayMetrics). */
    function vesselCapacityTotals(v) {
        var pct = Math.max(0, Math.min(100, Number(v.pctFull) || 0));
        var capH = v.capacityHeightFt != null ? Number(v.capacityHeightFt) : 14;
        if (isNaN(capH) || capH <= 0) capH = 14;
        var prodVol = Number(v.volumeCuFt) || 0;
        var capVol = pct > 0.05 ? prodVol / (pct / 100) : prodVol;
        var prodWt = Number(v.weightLb) || 0;
        var capWt = pct > 0.05 ? prodWt / (pct / 100) : prodWt;
        return { capH: capH, capVol: capVol, capWt: capWt };
    }

    function fmtFixed(n, d) {
        var x = Number(n);
        if (isNaN(x)) x = 0;
        return x.toFixed(d != null ? d : 2);
    }

    function vesselTypeShapeLabel(v) {
        ensureVesselSetupDefaults(v);
        var tid = parseInt(v.vesselTypeId, 10) || 1;
        var vt = VESSEL_TYPES.filter(function (x) {
            return x.id === tid;
        })[0];
        var desc = vt ? vt.name : "Vessel";
        var ps = parseFloat(v.partitionScale);
        if (v.verticalSplitPartitioned && !isNaN(ps) && ps < 100) {
            return fmtFixed(ps, 0) + "% Partition of " + desc;
        }
        return desc;
    }

    function vesselHasMeasurementRecord(v) {
        return !!(v.lastMeasurement && String(v.lastMeasurement).trim());
    }

    function vesselDetailsNextScheduleText(v) {
        var st = parseInt(v.sensorTypeId, 10);
        if (st === 14 || st === 15) return "Unknown Schedule";
        var vid = v.id;
        var applies = function (s) {
            ensureScheduleFields(s);
            if (s.eventType != null && s.eventType !== MEASUREMENT_EVENT_TYPE) return false;
            if (s.draft) return false;
            if (s.scheduleActive === false) return false;
            if (s.vesselIds && s.vesselIds.indexOf(vid) >= 0) return true;
            if (s.groupIds && s.groupIds.length) {
                var gi;
                for (gi = 0; gi < s.groupIds.length; gi++) {
                    var g = state.groups.filter(function (x) {
                        return x.id === s.groupIds[gi];
                    })[0];
                    if (g && g.vesselIds && g.vesselIds.indexOf(vid) >= 0) return true;
                }
            }
            return false;
        };
        var list = schedulesMeasurementList().filter(applies);
        if (list.length === 0) return "None Scheduled";
        var sch = list[0];
        var t = (sch.scheduleInfoTime || sch.time || "08:00").slice(0, 5);
        var d = sch.scheduleStartDate || todayIsoDateLocal();
        return d + " " + t;
    }

    /** Bottom of client area (Control color) — lblLastMeasure / lblNextMeasure; btnClose is separate on white bar below (Designer). */
    function buildVesselDetailsBottomScheduleHtml(v) {
        ensureVesselFields(v, 0);
        var hasMeas = vesselHasMeasurementRecord(v);
        var lastL = hasMeas ? v.lastMeasurement : "None Taken";
        var nextL = vesselDetailsNextScheduleText(v);
        return (
            '<div class="vd-bottom-sched" aria-label="Last and next measurement">' +
            '<span class="vd-sched-lbl">Last:</span>' +
            '<span class="vd-sched-val">' +
            escapeHtml(lastL) +
            "</span>" +
            '<span class="vd-sched-lbl">Next:</span>' +
            '<span class="vd-sched-val">' +
            escapeHtml(nextL) +
            "</span>" +
            "</div>"
        );
    }

    function buildVesselDetailsFooterHtml(v) {
        return (
            '<div class="vd-footer-inner">' +
            buildVesselDetailsBottomScheduleHtml(v) +
            '<button type="button" class="secondary vd-footer-close" data-close-app>Close</button>' +
            "</div>"
        );
    }

    function alarmDetailText(enabled, pctVal) {
        if (!enabled) return "OFF";
        var p = Number(pctVal);
        if (isNaN(p)) return "OFF";
        return fmtFixed(p, 0) + "%";
    }

    /** Stable pseudo-random ints/floats for simulated register dumps (per vessel + address). */
    function tutorialSensorSeed(v, salt) {
        var s = String(v && v.id ? v.id : "") + "\0" + String(v && v.sensorAddress != null ? v.sensorAddress : "") + "\0" + String(salt || "");
        var h = 5381;
        var i;
        for (i = 0; i < s.length; i++) {
            h = ((h << 5) + h + s.charCodeAt(i)) | 0;
        }
        return Math.abs(h);
    }

    function simInt(seed, idx, lo, hi) {
        var x = ((seed * (idx + 499)) ^ (seed >>> 7)) >>> 0;
        return lo + (x % (hi - lo + 1));
    }

    function simFloat(seed, idx, lo, hi) {
        var x = (((seed + idx * 9973) >>> 0) % 1000000) / 1000000;
        return lo + x * (hi - lo);
    }

    /**
     * Mirrors BobsBO Sensor*.RefreshSensorDetails → strSensorData / BobMsgQue SensorData.DataString.
     * LevelScanner.vb, SmartSonicWave.vb, MnuUltrasonic.vb, MpxMagnetostrictive.vb, Pt400500Pressure.vb,
     * SmartBobC100MB.vb, Spl100.vb, Spl200.vb, Ncr80Gwr2000.vb; HART abbreviated.
     */
    function sensorDetailsDataString(v, deviceAddrStr) {
        ensureVesselSetupDefaults(v);
        var st = parseInt(v.sensorTypeId, 10) || 4;
        var addr = deviceAddrStr != null ? String(deviceAddrStr).trim() : String(v.sensorAddress || "1").trim();
        var seed = tutorialSensorSeed(v, "dt:" + st + ":" + addr);
        var pct = Number(v.pctFull) || 0;
        var dp = parseInt(v.sensorDecimalPlaces, 10);
        if (isNaN(dp)) dp = 2;
        var distReg = st === 5 ? 324 : 320;

        if (st === 10) {
            /* LevelScanner.RefreshSensorDetails — labels/spacing match LevelScanner.vb */
            var mid = simFloat(seed, 1, 2.4, 4.2);
            var sngAvgDistance = mid;
            var sngMinDistance = mid - simFloat(seed, 2, 0.1, 0.6);
            var sngMaxDistance = mid + simFloat(seed, 3, 0.3, 1.1);
            var sngVolume = pct;
            var sngAnalogOut = 4 + pct / 25;
            var sngSnr = simFloat(seed, 4, 0, 5);
            var sngTemperatureCelcius = simFloat(seed, 5, -52, 45);
            var sngTemperatureFahrenheit = sngTemperatureCelcius * (9 / 5) + 32;
            function f6(x) {
                return Number(x).toFixed(6);
            }
            var lines = [
                "Average Distance (m):  " + f6(sngAvgDistance),
                "Minimum Distance (m):  " + f6(sngMinDistance),
                "Maximum Distance (m):  " + f6(sngMaxDistance),
                "Volume (%):            " + sngVolume,
                "Analog Out (mA):       " + sngAnalogOut,
                "SNR (dB):              " + sngSnr,
                "Temperature (ºC):      " + sngTemperatureCelcius,
                "Temperature (ºF):      " + sngTemperatureFahrenheit
            ];
            return lines.join("\r\n");
        }

        if (st === 4 || st === 5) {
            /* SmartSonicWave.RefreshSensorDetails */
            var linesB = [
                "Sensor ID:          " + addr,
                "Distance Register:  " + distReg,
                "Decimal Places:     " + dp
            ];
            return linesB.join("\r\n");
        }

        if (st === 6) {
            /* MnuUltrasonic — full label set, simulated register integers */
            var L = function (label, idx) {
                return label + simInt(seed, idx, 0, 9999);
            };
            var parts = [];
            parts.push("------- INPUT REGISTERS -------");
            parts.push(L("Model Type:                ", 10));
            parts.push(L("Raw Distance:              ", 11));
            parts.push(L("Temperature:               ", 12));
            parts.push(L("Calculated Reading:        ", 13));
            parts.push(L("Version:                   ", 14));
            parts.push(L("Signal Strength:           ", 15));
            parts.push(L("Trip 1 Alarm:              ", 16));
            parts.push(L("Trip 1 Status:             ", 17));
            parts.push(L("Trip 2 Alarm:              ", 18));
            parts.push(L("Trip 2 Status:             ", 19));
            parts.push("");
            parts.push("------ HOLDING REGISTERS ------");
            parts.push(L("Device Address:            ", 20));
            parts.push(L("Units:                     ", 21));
            parts.push(L("Application Type:          ", 22));
            parts.push(L("Volume Units:              ", 23));
            parts.push(L("Decimal Calculated:        ", 24));
            parts.push(L("Max Distance:              ", 25));
            parts.push(L("Full Distance:             ", 26));
            parts.push(L("Empty Distance:            ", 27));
            parts.push(L("Sensitivity:               ", 28));
            parts.push(L("Pulses:                    ", 29));
            parts.push(L("Blanking:                  ", 30));
            parts.push(L("Gain Control:              ", 31));
            parts.push(L("Averaging:                 ", 32));
            parts.push(L("Filter Window:             ", 33));
            parts.push(L("Out of Range:              ", 34));
            parts.push(L("Sample Rate:               ", 35));
            parts.push(L("Multiplier:                ", 36));
            parts.push(L("Offset:                    ", 37));
            parts.push(L("Temperature Compensation:  ", 38));
            parts.push(L("Trip 1 Value:              ", 39));
            parts.push(L("Trip 1 Window:             ", 40));
            parts.push(L("Trip 1 Type:               ", 41));
            parts.push(L("Trip 2 Value:              ", 42));
            parts.push(L("Trip 2 Window:             ", 43));
            parts.push(L("Trip 2 Type:               ", 44));
            parts.push(L("Parameter 1:               ", 45));
            parts.push(L("Parameter 2:               ", 46));
            parts.push(L("Parameter 3:               ", 47));
            parts.push(L("Parameter 4:               ", 48));
            parts.push(L("Parameter 5:               ", 49));
            return parts.join("\r\n");
        }

        if (st === 7) {
            var L7 = function (label, idx) {
                return label + simInt(seed, idx, 0, 9999);
            };
            var p7 = [];
            p7.push("------- INPUT REGISTERS -------");
            p7.push(L7("Raw Top Float:            ", 60));
            p7.push(L7("Raw Bottom Float:         ", 61));
            p7.push(L7("Temperature:              ", 62));
            p7.push(L7("Calculated Top Float:     ", 63));
            p7.push(L7("Calculated Bottom Float:  ", 64));
            p7.push(L7("Version:                  ", 65));
            p7.push("");
            p7.push("------ HOLDING REGISTERS ------");
            p7.push(L7("Device Address:           ", 66));
            p7.push(L7("Units:                    ", 67));
            p7.push(L7("Application Type:         ", 68));
            p7.push(L7("Volume Units:             ", 69));
            p7.push(L7("Decimal Place:            ", 70));
            p7.push(L7("Max Distance:             ", 71));
            p7.push(L7("Full Distance:            ", 72));
            p7.push(L7("Empty Distance:           ", 73));
            p7.push(L7("Sensitivity:              ", 74));
            p7.push(L7("Pulses:                   ", 75));
            p7.push(L7("Blanking:                 ", 76));
            p7.push(L7("Averaging:                ", 77));
            p7.push(L7("Filter Window:            ", 78));
            p7.push(L7("Out of Range Samples:     ", 79));
            p7.push(L7("Sample Rate:              ", 80));
            p7.push(L7("Multiplier:               ", 81));
            p7.push(L7("Offset:                   ", 82));
            p7.push(L7("Pre Filter:               ", 83));
            p7.push(L7("Noise Limit:              ", 84));
            p7.push(L7("Temperature Select:       ", 85));
            p7.push(L7("RTC Offset:               ", 86));
            p7.push(L7("Float Window:             ", 87));
            p7.push(L7("1st Float Offset:         ", 88));
            p7.push(L7("2nd Float Offset:         ", 89));
            p7.push(L7("Gain Offset:              ", 90));
            p7.push(L7("4mA Set Point:            ", 91));
            p7.push(L7("20mA Set Point:           ", 92));
            p7.push(L7("4mA Calibration:          ", 93));
            p7.push(L7("20mA Calibration:         ", 94));
            p7.push(L7("Web Alarm 1 Distance:     ", 95));
            p7.push(L7("Web Alarm 1 Window:       ", 96));
            p7.push(L7("Web Alarm 1 Type:         ", 97));
            p7.push(L7("Web Alarm 2 Distance:     ", 98));
            p7.push(L7("Web Alarm 2 Window:       ", 99));
            p7.push(L7("Web Alarm 2 Type:         ", 100));
            p7.push(L7("Parameter 1:              ", 101));
            p7.push(L7("Parameter 2:              ", 102));
            p7.push(L7("Parameter 3:              ", 103));
            p7.push(L7("Parameter 4:              ", 104));
            p7.push(L7("Parameter 5:              ", 105));
            return p7.join("\r\n");
        }

        if (st === 8 || st === 9) {
            /* Pt400500Pressure — model type 9 branch (level-style), simulated */
            var Lp = function (label, idx) {
                return label + simInt(seed, idx, 0, 999);
            };
            var Fp = function (label, idx) {
                return label + simFloat(seed, idx, -1, 1).toFixed(4);
            };
            var pp = [];
            pp.push("------- INPUT REGISTERS -------");
            pp.push(Lp("Model Type:          ", 200));
            pp.push(Lp("Level:               ", 201));
            pp.push(Lp("Temperature:         ", 202));
            pp.push(Lp("Calculated Reading:  ", 203));
            pp.push(Lp("Battery Voltage:     ", 204));
            pp.push(Lp("Trip 1 Status:       ", 205));
            pp.push(Lp("Trip 2 Status:       ", 206));
            pp.push("");
            pp.push("------ HOLDING REGISTERS ------");
            pp.push(Lp("Device Address:      ", 207));
            pp.push(Lp("Units:               ", 208));
            pp.push(Lp("Application Type:    ", 209));
            pp.push(Lp("Volume Units:        ", 210));
            pp.push(Lp("Decimal Calculated:  ", 211));
            pp.push(Lp("Max Level:           ", 212));
            pp.push(Lp("Full Level:          ", 213));
            pp.push(Lp("Zero Offset:         ", 214));
            pp.push(Lp("A/D Gain:            ", 215));
            pp.push(Lp("Specific Gravity:    ", 216));
            pp.push(Lp("Parameter Default:   ", 217));
            pp.push(Lp("Averaging:           ", 218));
            pp.push(Lp("Calibration Value:   ", 219));
            pp.push(Lp("Calibration Flag:    ", 220));
            pp.push(Lp("Sample Rate:         ", 221));
            pp.push(Lp("Scale:               ", 222));
            pp.push(Lp("Offset:              ", 223));
            pp.push(Lp("Voltage Offset:      ", 224));
            pp.push(Lp("Baud Rate:           ", 225));
            pp.push(Lp("Parity:              ", 226));
            pp.push(Lp("Stop Bit:            ", 227));
            pp.push(Fp("Pressure X3:         ", 228));
            pp.push(Fp("Pressure X2:         ", 229));
            pp.push(Fp("Pressure X1:         ", 230));
            pp.push(Fp("Pressure X0:         ", 231));
            pp.push(Lp("Trip 1 Level:        ", 232));
            pp.push(Lp("Trip 1 Window:       ", 233));
            pp.push(Lp("Trip 1 Type:         ", 234));
            pp.push(Lp("Trip 2 Level:        ", 235));
            pp.push(Lp("Trip 2 Window:       ", 236));
            pp.push(Lp("Trip 2 Type:         ", 237));
            pp.push(Lp("Parameter 1:         ", 238));
            pp.push(Lp("Parameter 2:         ", 239));
            pp.push(Lp("Parameter 3:         ", 240));
            pp.push(Lp("Parameter 4:         ", 241));
            pp.push(Lp("Parameter 5:         ", 242));
            pp.push(Lp("Temperature Offset:  ", 243));
            pp.push(Fp("Temperature X3:      ", 244));
            pp.push(Fp("Temperature X2:      ", 245));
            pp.push(Fp("Temperature X1:      ", 246));
            pp.push(Fp("Temperature X0:      ", 247));
            return pp.join("\r\n");
        }

        if (st === 13) {
            /* SmartBobC100MB.RefreshSensorDetails — idle, valid data */
            var drop = simInt(seed, 300, 1000, 5000);
            var linesC = [
                "Measurement Status:        Idle",
                "Results Status:            Valid Data",
                "SmartBob Status:           Enabled",
                "Drop Measurement:          " + drop + " (" + (drop / 80).toFixed(2) + " ft)",
                "C-100MB Calculated Level:  " + pct.toFixed(2) + "%"
            ];
            return linesC.join("\r\n");
        }

        if (st === 14) {
            var dist = simFloat(seed, 400, 1, 15).toFixed(3);
            var spl = [];
            spl.push("Address:             " + addr);
            spl.push("");
            spl.push("Distance (m):        " + dist);
            spl.push("Timestamp:           " + (v.lastMeasurement || ""));
            spl.push("Reading Count:       " + simInt(seed, 401, 1, 50));
            spl.push("");
            spl.push("Battery (V):         " + simFloat(seed, 402, 2.5, 3.6).toFixed(2));
            spl.push("Timestamp:           " + (v.lastMeasurement || ""));
            spl.push("");
            spl.push("Battery Loaded (V):  " + simFloat(seed, 403, 2.5, 3.6).toFixed(2));
            spl.push("Timestamp:           " + (v.lastMeasurement || ""));
            spl.push("");
            spl.push("SNR:                 " + simFloat(seed, 404, 5, 30).toFixed(1));
            return spl.join("\r\n");
        }

        if (st === 15) {
            var spl2 = [];
            spl2.push("Sensor Address:       " + addr);
            spl2.push("Distance (ft)/Error:  " + pct.toFixed(2));
            spl2.push("Reading Count:        " + simInt(seed, 500, 1, 200));
            spl2.push("MCU Temperature (C):  " + simFloat(seed, 501, 15, 40).toFixed(1));
            spl2.push("Battery (V):          " + simFloat(seed, 502, 2.5, 3.6).toFixed(2));
            spl2.push("Temperature (C):      " + simFloat(seed, 503, 10, 35).toFixed(1));
            spl2.push("RSSI (dB):            " + simInt(seed, 504, -90, -40));
            spl2.push("SNR (dB):             " + simFloat(seed, 505, 5, 25).toFixed(1));
            spl2.push("Packet Length:        " + simInt(seed, 506, 40, 120));
            return spl2.join("\r\n");
        }

        if (st === 16) {
            var hh = [];
            hh.push("Primary Variable:      " + pct.toFixed(2) + " %");
            hh.push("Loop Current (mA):     " + (4 + pct / 25).toFixed(2));
            hh.push("Upper Range:           100");
            hh.push("Lower Range:           0");
            hh.push("Damping:               " + simFloat(seed, 600, 0, 5).toFixed(1));
            hh.push("Device Status:         " + (v.status || "Ready"));
            hh.push("Extended Status:       0");
            return hh.join("\r\n");
        }

        if (st === 11 || st === 12) {
            /* Ncr80Gwr2000.RefreshSensorDetails — intStatusInvalidMeasurement: 0 = OK, 0xF = fault (tutorial mirrors comm/fault) */
            /*
             * Engineering unit codes match BobsBO Ncr80Gwr2000.vb + tblHartTable2 (Common Tables rev 25, table 2) / commented GetUnits.
             * GetDistanceInFeet: SV Unit 44 = SV in feet; 45 = meters; 47/48/49 = in/cm/mm.
             */
            var vegaMetric =
                state.systemSettings && String(state.systemSettings.units || "").toLowerCase() === "metric";
            var pctN = Math.max(0, Math.min(100, Number(v.pctFull) || 0));
            var pvJ = simFloat(seed, 700, -0.04, 0.04);
            var sngPV = Math.max(0, Math.min(100, pctN + pvJ));
            var sngTV = vegaMetric
                ? simFloat(seed, 702, -5, 42)
                : simFloat(seed, 702, 23, 108);
            var sngQV = vegaMetric
                ? simFloat(seed, 703, 0, 85)
                : simFloat(seed, 703, 0, 3000);
            /*
             * SV — distance (headspace / headroom), same geometry as Measure; long fractional digits match register dumps.
             */
            var mVega = vesselDisplayMetrics(v);
            var headFt =
                mVega.headroom && mVega.headroom.heightFt != null ? Number(mVega.headroom.heightFt) : 0;
            if (isNaN(headFt)) headFt = 0;
            /* 44 = feet, 45 = meters (same as Ncr80Gwr2000.GetDistanceInFeet — not generic HART-only aliases) */
            var svUnitCode = vegaMetric ? 45 : 44;
            var baseHeadDist = vegaMetric ? headFt * 0.3048 : headFt;
            var svJitter = simFloat(seed, 701, -0.0012, 0.0012);
            var sngSVRaw = baseHeadDist + svJitter;
            var sngSVStr = Number(sngSVRaw).toFixed(11);
            /* 57 = percent (HART table II — primary variable level / %) */
            var pvUnitCode = 57;
            /* 32 = °C, 33 = °F (commented GetUnits in Ncr80Gwr2000.vb) */
            var tvUnitCode = vegaMetric ? 32 : 33;
            /* 43 = m³, 112 = ft³ — quaternary often volume/derived; scales with site */
            var qvUnitCode = vegaMetric ? 43 : 112;
            var vega = [];
            vega.push("----- INPUT REGISTERS -----");
            vega.push("Status:          " + ncrGwrInvalidMeasurementStatusDisplay(v));
            vega.push("PV Unit:         " + pvUnitCode);
            vega.push("PV:              " + sngPV);
            vega.push("SV Unit:         " + svUnitCode);
            vega.push("SV:              " + sngSVStr);
            vega.push("TV Unit:         " + tvUnitCode);
            vega.push("TV:              " + sngTV);
            vega.push("QV Unit:         " + qvUnitCode);
            vega.push("QV:              " + sngQV);
            vega.push("");
            vega.push("---- HOLDING REGISTERS ----");
            vega.push("Device Address:  " + addr);
            vega.push("Baud Rate:       " + simInt(seed, 709, 0, 5));
            vega.push("Parity:          " + simInt(seed, 710, 0, 2));
            vega.push("Stop Bits:       " + simInt(seed, 711, 1, 2));
            vega.push("Delay Time:      " + simInt(seed, 712, 0, 500));
            return vega.join("\r\n");
        }

        return (
            "SensorTypeID=" +
            st +
            " DeviceAddress=" +
            addr +
            "\r\nStatus: " +
            (v.status || defaultIdleStatusForSensorType(st))
        );
    }

    /**
     * NCR-80 / GWR-2000 input register invalid-measurement / status nibble — working = 0x0, fault = 0xF (BobsBO Ncr80Gwr2000).
     */
    function ncrGwrInvalidMeasurementStatusDisplay(v) {
        if (!v) return "0xF";
        if (v.sensorEnabled === false) return "0xF";
        if (state.vesselsReadOnly || !state.ebobServicesRunning) return "0xF";
        var s = (v.status || "").toString().trim();
        if (s === "Communication Error") return "0xF";
        if (/connection lost/i.test(s)) return "0xF";
        if (s === "Motor Fault") return "0xF";
        if (s.indexOf("Stuck") >= 0) return "0xF";
        if (s === "Measurement Error") return "0xF";
        if (s.indexOf("Communication Error") >= 0) return "0xF";
        /* Idle / Ready / Unknown / Pending / in-cycle strings → healthy register nibble */
        return "0x0";
    }

    /**
     * tblDevice rows for Protocol A — single row from vessel fields, or v.sbDevices[] for multi-bob.
     * Sorted by address ascending (same ORDER BY as eBob SQL).
     */
    function smartBobADeviceRows(v) {
        ensureVesselSetupDefaults(v);
        if (v.sbDevices && v.sbDevices.length > 0) {
            return v.sbDevices
                .map(function (d) {
                    var rw = parseFloat(d.relativeWeight != null ? d.relativeWeight : 1);
                    return {
                        address: d.address != null ? String(d.address).trim() : "",
                        label: d.label != null ? String(d.label) : "",
                        isActive: d.isActive !== false,
                        relativeFraction: isNaN(rw) ? 1 : rw
                    };
                })
                .filter(function (r) {
                    return r.address !== "";
                })
                .sort(function (a, b) {
                    return parseInt(a.address, 10) - parseInt(b.address, 10);
                });
        }
        var b = getSensorAddressBounds(v.sensorTypeId);
        var vr = validateIntegerSensorAddress(v.sensorAddress, b.min, b.max);
        var addrStr = vr.ok ? vr.normalized : String(v.sensorAddress || "").trim();
        var rwf = parseFloat(v.deviceRelativeWeight != null ? v.deviceRelativeWeight : 1);
        if (isNaN(rwf)) rwf = 1;
        return [
            {
                address: addrStr,
                label: v.deviceLabel != null ? String(v.deviceLabel) : "",
                isActive: v.sensorEnabled !== false,
                relativeFraction: rwf
            }
        ];
    }

    /** tblMeasurementData_Temp LEFT JOIN — empty until temp row exists (RefreshSmartBobA). */
    function getSmartBobAHeadroomProductForDetails(v, addrStr, isActive) {
        if (!isActive) return { head: "", prod: "" };
        ensureVesselSetupDefaults(v);
        var a = String(addrStr).trim();
        var t = v.sbMeasurementTempByAddr && v.sbMeasurementTempByAddr[a];
        if (!t) return { head: "", prod: "" };
        var abbr = t.distanceAbbrev || "ft";
        var hh = t.headroomHeight;
        var pp = t.productHeight;
        if (hh == null || isNaN(hh) || pp == null || isNaN(pp)) return { head: "", prod: "" };
        return {
            head: fmtFixed(hh, 2) + " " + abbr,
            prod: fmtFixed(pp, 2) + " " + abbr
        };
    }

    /** frmVesselDetails InitializeSensorSmartBobA — Format(RelativeWeight * 100, "##0.00") & "%". */
    function formatSmartBobARelativeWeightDisplay(row) {
        if (!row || !row.isActive) return "";
        var f = row.relativeFraction;
        if (f == null || isNaN(f)) f = 1;
        return fmtFixed(Number(f) * 100, 2) + "%";
    }

    function buildVesselDetailsSensorsPanelHtml(v) {
        ensureVesselFields(v, 0);
        var st = parseInt(v.sensorTypeId, 10) || 1;
        var net = getSensorNetworkById(v.sensorNetworkId);
        var netLine = formatNetworkSelectLabel(net);
        var dt = deviceTypeById(st);
        var typeLine = dt ? dt.name : "Unknown";

        var headNet =
            '<div class="vd-sens-head">' +
            '<div class="vd-sens-line"><span class="vd-sens-lbl">Sensor Network:</span> ' +
            escapeHtml(netLine) +
            "</div>" +
            '<div class="vd-sens-line vd-sens-line-type"><span class="vd-sens-lbl">Sensor Type:</span> ' +
            escapeHtml(typeLine) +
            "</div></div>";

        if (st === 1 || st === 3) {
            var rowsHtml = smartBobADeviceRows(v)
                .map(function (row) {
                    var stat = row.isActive ? v.status || "Unknown" : "Disabled";
                    var hp = getSmartBobAHeadroomProductForDetails(v, row.address, row.isActive);
                    var rwDisp = formatSmartBobARelativeWeightDisplay(row);
                    return (
                        '<div class="vd-sb-a-lv-row vd-sb-a-lv-data" role="row">' +
                        '<span class="vd-sb-a-c-addr" role="cell">' +
                        escapeHtml(row.address) +
                        "</span>" +
                        '<span class="vd-sb-a-c-label" role="cell">' +
                        escapeHtml(row.label || "") +
                        "</span>" +
                        '<span class="vd-sb-a-c-status" role="cell">' +
                        escapeHtml(stat) +
                        "</span>" +
                        '<span class="vd-sb-a-c-hr" role="cell">' +
                        escapeHtml(hp.head) +
                        "</span>" +
                        '<span class="vd-sb-a-c-prod" role="cell">' +
                        escapeHtml(hp.prod) +
                        "</span>" +
                        '<span class="vd-sb-a-c-rw" role="cell">' +
                        escapeHtml(rwDisp) +
                        "</span></div>"
                    );
                })
                .join("");
            return (
                headNet +
                '<div class="vd-sb-a-wrap"><div class="vd-sb-a-lv" role="table">' +
                '<div class="vd-sb-a-lv-row vd-sb-a-lv-hdr" role="row">' +
                '<span class="vd-sb-a-c-addr vd-sb-a-hu" role="columnheader">Addr</span>' +
                '<span class="vd-sb-a-c-label vd-sb-a-hu" role="columnheader">Label</span>' +
                '<span class="vd-sb-a-c-status vd-sb-a-hu" role="columnheader">Status</span>' +
                '<span class="vd-sb-a-c-hr vd-sb-a-hu" role="columnheader">Headroom</span>' +
                '<span class="vd-sb-a-c-prod vd-sb-a-hu" role="columnheader">Product</span>' +
                '<span class="vd-sb-a-c-rw vd-sb-a-hu vd-sb-a-hdr-rw" role="columnheader" aria-label="Relative Weight">' +
                '<span class="vd-sb-a-hdr-rw-lines" aria-hidden="true">' +
                '<span class="vd-sb-a-hdr-rw-line">Relative</span>' +
                '<span class="vd-sb-a-hdr-rw-line">Weight</span></span></span>' +
                "</div>" +
                (rowsHtml || '<div class="vd-sb-a-empty-wrap"><span class="vd-sb-empty">No devices.</span></div>') +
                "</div></div>"
            );
        }

        if (st === 2) {
            var addrB = String(v.sensorAddress || "1").trim();
            var stB = v.sensorEnabled !== false ? v.status || "Unknown" : "Disabled";
            var dc = v.smartBobDropCount != null ? String(v.smartBobDropCount) : "";
            var rc = v.smartBobRetractCount != null ? String(v.smartBobRetractCount) : "";
            return (
                headNet +
                '<div class="vd-sb-b-grid">' +
                "<div><span class=\"vd-k\">Address:</span> " +
                escapeHtml(addrB) +
                "</div>" +
                "<div><span class=\"vd-k\">Status:</span> " +
                escapeHtml(stB) +
                "</div>" +
                "<div><span class=\"vd-k\">Drop Count:</span> " +
                escapeHtml(dc) +
                "</div>" +
                "<div><span class=\"vd-k\">Retract Count:</span> " +
                escapeHtml(rc) +
                "</div></div>"
            );
        }

        if (st === 4 || st === 5 || st === 6 || st === 7 || st === 8 || st === 9 || st === 10 || st === 13 || st === 14 || st === 15 || st === 16) {
            var raw = sensorDetailsDataString(v, v.sensorAddress);
            return (
                headNet +
                '<div class="vd-sensor-block"><pre class="vd-sensor-raw" spellcheck="false">' +
                escapeHtml(raw) +
                "</pre></div>"
            );
        }

        if (st === 11 || st === 12) {
            ensureVegaSensorsArray(v);
            var n = Math.min(32, Math.max(1, parseInt(v.vegaSensorCount || "1", 10) || 1));
            var m2 = vesselDisplayMetrics(v);
            var du2 = "ft";
            var status0 = v.sensorEnabled !== false ? v.status || "Unknown" : "Disabled";
            var sumRows = [];
            var i;
            var innerTabs = [];
            var nEn = 0;
            for (i = 0; i < n; i++) {
                if (v.vegaSensors[i] && v.vegaSensors[i].enabled) nEn += 1;
            }
            if (nEn < 1) nEn = 1;
            var rwEach = 100 / nEn;
            for (i = 0; i < n; i++) {
                var row = v.vegaSensors[i];
                var en = row && row.enabled;
                var adr = row && row.address != null ? String(row.address).trim() : String(i + 1);
                var name = "Address " + adr;
                if (row && row.label && String(row.label).trim() !== "") {
                    name += " [" + String(row.label).trim() + "]";
                }
                var stx = !en ? "Disabled" : status0;
                var hr2 = en ? fmtFixed(m2.headroom.heightFt, 2) + " " + du2 : "";
                var pr2 = en ? fmtFixed(m2.product.heightFt, 2) + " " + du2 : "";
                var rw2 = en ? fmtFixed(rwEach, 2) + "%" : "";
                var lm = en ? v.lastMeasurement || "" : "";
                sumRows.push(
                    "<tr>" +
                        '<td class="vd-vega-td-sensor">' +
                        escapeHtml(name) +
                        "</td>" +
                        '<td class="vd-vega-td-status">' +
                        escapeHtml(stx) +
                        "</td>" +
                        '<td class="vd-vega-td-hr">' +
                        escapeHtml(hr2) +
                        "</td>" +
                        '<td class="vd-vega-td-prod">' +
                        escapeHtml(pr2) +
                        "</td>" +
                        '<td class="vd-vega-td-wt">' +
                        escapeHtml(rw2) +
                        "</td>" +
                        '<td class="vd-vega-td-dt">' +
                        escapeHtml(lm) +
                        "</td></tr>"
                );
                var rawV = sensorDetailsDataString(v, adr);
                innerTabs.push({
                    name: name,
                    addr: adr,
                    body:
                        '<div class="vd-sensor-block"><pre class="vd-sensor-raw vd-vega-addr-pre" spellcheck="false">' +
                        escapeHtml(rawV) +
                        "</pre></div>"
                });
            }
            /*
             * frmVesselDetails.Designer.vb — lvwVegaSummary Size 490×201; default column widths:
             *   Sensor 192, Status 109, Headroom 62, Product 62, Weight 53, Date/Time 135 (sum 613).
             * frmVesselDetails.vb InitializeSensorVega: after each row,
             *   lvwVegaSummary.Columns.Item(0).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
             * so column 0 shrinks to header + longest "Address …" cell — not 192px.
             * Remaining columns keep Designer proportions; slack goes to Date/Time when col0 + 421 ≤ 490.
             */
            var VEGA_LIST_INNER_PX = 490;
            var W_ST = 109;
            var W_HR = 62;
            var W_PR = 62;
            var W_WT = 53;
            var W_DT = 135;
            var vegaFixedRest = W_ST + W_HR + W_PR + W_WT + W_DT;
            var longestVegaSensor = "Sensor";
            var vi;
            for (vi = 0; vi < sumRows.length; vi++) {
                /* Rebuild display name same as row loop (for AutoResize width) */
                var vr = v.vegaSensors[vi];
                var vadr = vr && vr.address != null ? String(vr.address).trim() : String(vi + 1);
                var vnm = "Address " + vadr;
                if (vr && vr.label && String(vr.label).trim() !== "") {
                    vnm += " [" + String(vr.label).trim() + "]";
                }
                if (vnm.length > longestVegaSensor.length) longestVegaSensor = vnm;
            }
            var hdrSens = "Sensor";
            var vegaCh = Math.max(hdrSens.length, longestVegaSensor.length);
            /* Match ListView ColumnHeaderAutoResizeStyle.ColumnContent (header + widest subitem) */
            var col0Px = Math.ceil(vegaCh * 6.1) + 14;
            col0Px = Math.max(52, Math.min(192, col0Px));
            var VEGA_COL_RAW;
            if (col0Px + vegaFixedRest > VEGA_LIST_INNER_PX) {
                var sc = (VEGA_LIST_INNER_PX - col0Px) / vegaFixedRest;
                VEGA_COL_RAW = [
                    col0Px,
                    Math.round(W_ST * sc),
                    Math.round(W_HR * sc),
                    Math.round(W_PR * sc),
                    Math.round(W_WT * sc),
                    Math.round(W_DT * sc)
                ];
            } else {
                var vegaSlack = VEGA_LIST_INNER_PX - col0Px - vegaFixedRest;
                VEGA_COL_RAW = [col0Px, W_ST, W_HR, W_PR, W_WT, W_DT + vegaSlack];
            }
            var vegaSumPx = VEGA_COL_RAW.reduce(function (a, b) {
                return a + b;
            }, 0);
            if (vegaSumPx !== VEGA_LIST_INNER_PX) {
                VEGA_COL_RAW[5] += VEGA_LIST_INNER_PX - vegaSumPx;
            }
            var colgroupHtml =
                "<colgroup>" +
                VEGA_COL_RAW.map(function (w) {
                    var pct = (100 * w) / VEGA_LIST_INNER_PX;
                    return '<col style="width:' + pct.toFixed(5) + '%" />';
                }).join("") +
                "</colgroup>";
            var vegaPad = "";
            var piPad;
            var dataRowCount = sumRows.length > 0 ? sumRows.length : 1;
            /* Pad rows fill ListView body without exceeding 201px — ~12px/row + header ~16px */
            var padCount = Math.max(0, 14 - dataRowCount);
            for (piPad = 0; piPad < padCount; piPad++) {
                vegaPad +=
                    '<tr class="vd-vega-sum-pad" aria-hidden="true"><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>';
            }
            var summaryTable =
                '<div class="vd-vega-summary">' +
                '<div class="vd-vega-sum-wrap">' +
                '<table class="vd-vega-sum-table" role="grid" aria-label="Sensor summary">' +
                colgroupHtml +
                "<thead><tr>" +
                '<th scope="col" class="vd-vega-th-sensor">Sensor</th>' +
                '<th scope="col" class="vd-vega-th-status">Status</th>' +
                '<th scope="col" class="vd-vega-th-mid">Headroom</th>' +
                '<th scope="col" class="vd-vega-th-mid">Product</th>' +
                '<th scope="col" class="vd-vega-th-mid">Weight</th>' +
                '<th scope="col" class="vd-vega-th-mid">Date/Time</th>' +
                "</tr></thead><tbody>" +
                (sumRows.join("") || "<tr><td colspan=\"6\" class=\"vd-vega-sum-empty\">No sensors.</td></tr>") +
                vegaPad +
                "</tbody></table></div></div>";
            var subNav =
                '<div class="vd-vega-subtabs" role="tablist">' +
                '<button type="button" class="vd-vega-subtab active" data-vd-vega-sub="sum" role="tab">Summary</button>' +
                innerTabs
                    .map(function (t, j) {
                        return (
                            '<button type="button" class="vd-vega-subtab" data-vd-vega-sub="' +
                            escapeHtml(String(j)) +
                            '" role="tab">' +
                            escapeHtml(t.name) +
                            "</button>"
                        );
                    })
                    .join("") +
                "</div>";
            var subPanels =
                '<div class="vd-vega-subpanel" data-vd-vega-panel="sum">' +
                summaryTable +
                "</div>" +
                innerTabs
                    .map(function (t, j) {
                        return (
                            '<div class="vd-vega-subpanel vd-hidden" data-vd-vega-panel="' +
                            j +
                            '">' +
                            t.body +
                            "</div>"
                        );
                    })
                    .join("");
            return (
                '<div class="vd-vega-sensors-root">' + headNet + subNav + subPanels + "</div>"
            );
        }

        return (
            headNet +
            '<p class="vd-sensor-na">Sensor details not available.</p>'
        );
    }

    function bindVesselDetailsModalInteractions() {
        var body = document.getElementById("appModalBody");
        if (!body || !body.querySelector(".vd-shell")) return;

        body.querySelectorAll("[data-vd-tab]").forEach(function (btn) {
            btn.addEventListener("click", function () {
                var id = btn.getAttribute("data-vd-tab");
                body.querySelectorAll("[data-vd-tab]").forEach(function (b) {
                    var on = b.getAttribute("data-vd-tab") === id;
                    b.setAttribute("aria-selected", on ? "true" : "false");
                    b.classList.toggle("active", on);
                });
                body.querySelectorAll("[data-vd-panel]").forEach(function (p) {
                    p.classList.toggle("vd-hidden", p.getAttribute("data-vd-panel") !== id);
                });
            });
        });

        body.querySelectorAll("[data-vd-vega-sub]").forEach(function (btn) {
            btn.addEventListener("click", function () {
                var id = btn.getAttribute("data-vd-vega-sub");
                body.querySelectorAll("[data-vd-vega-sub]").forEach(function (b) {
                    b.classList.toggle("active", b.getAttribute("data-vd-vega-sub") === id);
                });
                body.querySelectorAll("[data-vd-vega-panel]").forEach(function (p) {
                    p.classList.toggle("vd-hidden", p.getAttribute("data-vd-vega-panel") !== id);
                });
            });
        });
    }

    function buildVesselDetailsBodyHtml(v) {
        ensureVesselFields(v, 0);
        var hasMeas = vesselHasMeasurementRecord(v);
        var pct = Math.round(Number(v.pctFull) || 0);
        var m = vesselDisplayMetrics(v);
        var cap = vesselCapacityTotals(v);
        var contentsDisp =
            v.contents != null && String(v.contents).trim() !== ""
                ? String(v.contents).trim()
                : v.product != null && String(v.product).trim() !== ""
                  ? String(v.product).trim()
                  : "";
        var nameDisp = v.name != null && String(v.name).trim() !== "" ? String(v.name).trim() : "";
        var pctLabel = hasMeas ? pct + "% Full" : "% Full";
        var capHStr = hasMeas ? fmtFixed(cap.capH, 2) + " ft" : "";
        var capPair = hasMeas ? formatVolumeWeightPair(v, cap.capVol, cap.capWt) : { volStr: "", wtStr: "" };
        var capVStr = hasMeas ? capPair.volStr : "";
        var capWStr = hasMeas ? capPair.wtStr : "";

        var tid = parseInt(v.vesselTypeId, 10) || 1;
        var densityBlock = "";
        if (tid === 15) {
            densityBlock = "";
        } else if (v.densityMode === "sg") {
            var sgShow = "";
            if (v.specificGravity != null && String(v.specificGravity).trim() !== "") {
                var sgN = Number(v.specificGravity);
                if (!isNaN(sgN)) sgShow = fmtFixed(sgN, 4);
            }
            densityBlock =
                '<div class="vd-kv-row"><span class="vd-k">Specific Gravity:</span><span class="vd-v">' +
                escapeHtml(sgShow) +
                "</span></div>";
        } else {
            densityBlock =
                '<div class="vd-kv-row"><span class="vd-k">Product Density:</span><span class="vd-v">' +
                escapeHtml(fmtFixed(Number(v.productDensity) || 0, 2)) +
                " " +
                escapeHtml((v.densityUnits || "lbs / cubic ft").replace(/\s*\/\s*/g, " / ")) +
                "</span></div>";
        }

        var prodH = hasMeas ? fmtFixed(m.product.heightFt, 2) + " ft" : "";
        var prodPair = hasMeas
            ? formatVolumeWeightPair(v, Number(v.volumeCuFt) || 0, Number(v.weightLb) || 0)
            : { volStr: "", wtStr: "" };
        var prodV = hasMeas ? prodPair.volStr : "";
        var prodW = hasMeas ? prodPair.wtStr : "";
        var headH = hasMeas ? fmtFixed(m.headroom.heightFt, 2) + " ft" : "";
        var headPair = hasMeas
            ? formatVolumeWeightPair(v, m.headroom.volumeCuFt, m.headroom.weightLb)
            : { volStr: "", wtStr: "" };
        var headV = hasMeas ? headPair.volStr : "";
        var headW = hasMeas ? headPair.wtStr : "";

        var vesselTabInner =
            '<div class="vd-groupbox">' +
            '<div class="vd-groupbox-cap">Vessel Information</div>' +
            '<div class="vd-groupbox-inner">' +
            '<div class="vd-kv-row"><span class="vd-k">Vessel Name:</span><span class="vd-v">' +
            escapeHtml(nameDisp) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Type/Shape:</span><span class="vd-v">' +
            escapeHtml(vesselTypeShapeLabel(v)) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Capacity Height:</span><span class="vd-v">' +
            escapeHtml(capHStr) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Capacity Volume:</span><span class="vd-v">' +
            escapeHtml(capVStr) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Capacity Weight:</span><span class="vd-v">' +
            escapeHtml(capWStr) +
            "</span></div></div></div>" +
            '<div class="vd-groupbox">' +
            '<div class="vd-groupbox-cap">Content Information</div>' +
            '<div class="vd-groupbox-inner">' +
            '<div class="vd-cnt-top">' +
            '<div class="vd-kv-row"><span class="vd-k">Contents:</span><span class="vd-v">' +
            escapeHtml(contentsDisp) +
            "</span></div>" +
            densityBlock +
            "</div>" +
            '<div class="vd-cnt-grid">' +
            '<div class="vd-cnt-col">' +
            '<div class="vd-kv-row"><span class="vd-k">Content Height:</span><span class="vd-v">' +
            escapeHtml(prodH) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Content Volume:</span><span class="vd-v">' +
            escapeHtml(prodV) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Content Weight:</span><span class="vd-v">' +
            escapeHtml(prodW) +
            "</span></div></div>" +
            '<div class="vd-cnt-col vd-cnt-col-hr">' +
            '<div class="vd-kv-row"><span class="vd-k">Headroom Height:</span><span class="vd-v">' +
            escapeHtml(headH) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Headroom Volume:</span><span class="vd-v">' +
            escapeHtml(headV) +
            "</span></div>" +
            '<div class="vd-kv-row"><span class="vd-k">Headroom Weight:</span><span class="vd-v">' +
            escapeHtml(headW) +
            "</span></div></div></div></div></div>" +
            '<div class="vd-groupbox">' +
            '<div class="vd-groupbox-cap">Alarm Settings</div>' +
            '<div class="vd-groupbox-inner vd-alarm-grid">' +
            '<div class="vd-alarm-pair"><span class="vd-k">High Alarm:</span> ' +
            escapeHtml(alarmDetailText(v.alarmHighEnabled, v.alarmHighPct)) +
            "</div>" +
            '<div class="vd-alarm-pair"><span class="vd-k">Pre-Low Alarm:</span> ' +
            escapeHtml(alarmDetailText(v.alarmPreLowEnabled, v.alarmPreLowPct)) +
            "</div>" +
            '<div class="vd-alarm-pair"><span class="vd-k">Pre-High Alarm:</span> ' +
            escapeHtml(alarmDetailText(v.alarmPreHighEnabled, v.alarmPreHighPct)) +
            "</div>" +
            '<div class="vd-alarm-pair"><span class="vd-k">Low Alarm:</span> ' +
            escapeHtml(alarmDetailText(v.alarmLowEnabled, v.alarmLowPct)) +
            "</div></div></div>";

        var sensorsInner = buildVesselDetailsSensorsPanelHtml(v);

        return (
            '<div class="vd-shell">' +
            '<div class="vd-page-title">Vessel Details</div>' +
            '<div class="vd-main">' +
            '<div class="vd-left">' +
            '<div class="vd-contents">' +
            escapeHtml(contentsDisp) +
            "</div>" +
            '<div class="vd-vname">' +
            escapeHtml(nameDisp) +
            "</div>" +
            '<div class="vd-silo-preview" style="--vd-pct:' +
            (hasMeas ? pct : 0) +
            "%;" +
            escapeHtml(siloFillCssVarsFromColorName(v.fillColor || "Dark Red")) +
            '">' +
            '<div class="vd-silo-sky"></div><div class="vd-silo-fill"></div></div>' +
            '<div class="vd-pct">' +
            escapeHtml(pctLabel) +
            "</div></div>" +
            '<div class="vd-right">' +
            '<div class="vd-tabstrip" role="tablist">' +
            '<button type="button" class="vd-tab active" data-vd-tab="vessel" role="tab" aria-selected="true">Vessel</button>' +
            '<button type="button" class="vd-tab" data-vd-tab="sensors" role="tab" aria-selected="false">Sensors</button>' +
            "</div>" +
            '<div class="vd-panel" data-vd-panel="vessel">' +
            vesselTabInner +
            "</div>" +
            '<div class="vd-panel vd-hidden" data-vd-panel="sensors">' +
            sensorsInner +
            "</div></div></div>" +
            "</div>"
        );
    }

    function openVesselDetails(v) {
        if (!v) return;
        ensureVesselFields(v, 0);
        openAppModal(
            "Vessel Details - Binventory Workstation",
            buildVesselDetailsBodyHtml(v),
            buildVesselDetailsFooterHtml(v),
            "modal-vessel-details"
        );
        bindVesselDetailsModalInteractions();
    }

    function heightUnitsLabel() {
        return state.systemSettings && state.systemSettings.units === "Metric" ? "Meters (m)" : "Feet (ft)";
    }

    function fillColorPreviewCss(name) {
        var hx = FILL_COLOR_HEX_MAP[name] || "#cccccc";
        return "linear-gradient(180deg, " + hx + " 0%, #ffffff 100%)";
    }

    /** Parse #rrggbb for silo gradients (matches --silo-fill-* on :root, tinted by chosen fill color). */
    function hexToRgbForSilo(hex) {
        var h = String(hex || "").trim();
        var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(h);
        if (!m) return { r: 139, g: 0, b: 0 };
        return { r: parseInt(m[1], 16), g: parseInt(m[2], 16), b: parseInt(m[3], 16) };
    }

    function rgbCss(o) {
        return "rgb(" + o.r + "," + o.g + "," + o.b + ")";
    }

    function rgbAdjust(o, mult) {
        function c(x) {
            return Math.round(Math.min(255, Math.max(0, x)));
        }
        return { r: c(o.r * mult), g: c(o.g * mult), b: c(o.b * mult) };
    }

    /**
     * Inline style fragment setting --silo-fill-highlight and --silo-fill-depth for .silo / .vd-silo-preview.
     * Keeps the eBob two-layer look while following the vessel Fill Color choice.
     */
    function siloFillCssVarsFromColorName(name) {
        var hx = FILL_COLOR_HEX_MAP[name] || "#8B0000";
        var base = hexToRgbForSilo(hx);
        var top = rgbCss(rgbAdjust(base, 1.22));
        var mid = rgbCss(base);
        var bot = rgbCss(rgbAdjust(base, 0.7));
        var depth = "linear-gradient(180deg," + top + " 0%," + mid + " 52%," + bot + " 100%)";
        var br = base.r;
        var bg = base.g;
        var bb = base.b;
        var highlight =
            "linear-gradient(90deg," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.38) 0%," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.22) 14%," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.12) 32%," +
            "rgba(255,255,255,0.38) 49%," +
            "rgba(255,255,255,0.55) 50%," +
            "rgba(255,255,255,0.38) 51%," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.12) 68%," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.22) 86%," +
            "rgba(" +
            br +
            "," +
            bg +
            "," +
            bb +
            ",0.38) 100%)";
        return "--silo-fill-highlight:" + highlight + ";--silo-fill-depth:" + depth;
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

    function readTutorialMode() {
        var urlMode = "";
        try {
            var sp = new URLSearchParams(window.location.search || "");
            urlMode = String(sp.get("tutorial") || "").trim().toLowerCase();
        } catch (e) {
            urlMode = "";
        }
        if (urlMode) {
            try {
                sessionStorage.setItem(EBOB_TUTORIAL_MODE_KEY, urlMode);
            } catch (e2) {
                /* private mode */
            }
            return urlMode;
        }
        try {
            return String(sessionStorage.getItem(EBOB_TUTORIAL_MODE_KEY) || "").trim().toLowerCase();
        } catch (e3) {
            return "";
        }
    }

    function clearTutorialScenarioState() {
        try {
            sessionStorage.removeItem(EBOB_TUTORIAL_STATE_KEY);
        } catch (e) {
            /* private mode */
        }
    }

    function wireScenarioGateIfNeeded() {
        var gate = document.getElementById("tutorialScenarioGate");
        if (!gate) return false;
        var sp;
        try {
            sp = new URLSearchParams(window.location.search || "");
        } catch (e) {
            sp = null;
        }
        var modeFromQuery = sp ? String(sp.get("tutorial") || "").trim().toLowerCase() : "";
        if (modeFromQuery) {
            document.documentElement.classList.remove("ebob-initial-home-gate");
            gate.hidden = true;
            gate.setAttribute("aria-hidden", "true");
            document.body.classList.remove("ebob-scenario-gate-open");
            return false;
        }
        document.body.classList.add("ebob-scenario-gate-open");
        gate.hidden = false;
        gate.setAttribute("aria-hidden", "false");
        document.documentElement.classList.remove("ebob-initial-home-gate");
        if (gate.__ebobScenarioGateBound) return true;
        gate.__ebobScenarioGateBound = true;
        gate.addEventListener("click", function (e) {
            var btn = e.target.closest && e.target.closest("[data-scenario-mode]");
            if (!btn) return;
            e.preventDefault();
            var mode = String(btn.getAttribute("data-scenario-mode") || "").trim().toLowerCase();
            var implemented =
                mode === EBOB_TUTORIAL_MODE_DB_READ_ONLY ||
                mode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN ||
                mode === EBOB_TUTORIAL_MODE_FREE;
            if (!implemented) {
                alert("This tutorial scenario is not built yet.");
                return;
            }
            try {
                sessionStorage.setItem(EBOB_TUTORIAL_MODE_KEY, mode);
            } catch (err) {
                /* private mode */
            }
            var base = window.location.pathname || "ebob.html";
            window.location.href = base + "?tutorial=" + encodeURIComponent(mode);
        });
        return true;
    }

    function applyDatabaseReadOnlyTutorialScenario() {
        state.sites = state.sites.filter(function (s) {
            return s.id === "st1";
        });
        if (!state.sites.length) {
            state.sites = [
                {
                    id: "st1",
                    name: "Current Location",
                    workstationSiteId: 1,
                    serviceHostIp: "10.66.207.30",
                    serviceHostPort: "8093",
                    companyName: "Acme Manufacturing",
                    streetAddress: "100 Industrial Pkwy",
                    streetAddress2: "",
                    city: "Minneapolis",
                    state: "MN",
                    zip: "55401",
                    country: "USA",
                    distanceUnitsId: "1"
                }
            ];
        }
        ensureSiteFields(state.sites[0]);
        state.sites[0].workstationSiteId = 1;
        state.sites[0].serviceHostIp = "10.66.207.30";
        state.currentWorkstationSiteId = "st1";
        state.sensorNetworks = (state.sensorNetworks || []).filter(function (n) {
            return !n.siteId || n.siteId === "st1";
        });
        state.groups = (state.groups || []).filter(function (g) {
            return !g.siteId || g.siteId === "st1";
        });
        state.vessels = (state.vessels || []).filter(function (v) {
            return !v.siteId || v.siteId === "st1";
        });
        state.schedules = (state.schedules || []).filter(function (s) {
            return !s.siteId || s.siteId === "st1";
        });
    }

    /**
     * Pending/Unknown tutorial must use Modbus/RTU + NCR-80. Protocol A only allows SmartBob device types;
     * ensureVesselFields would reset sensorTypeId 11 → 1, breaking the glitch demo (SmartBob “retracting”) and Next gating.
     */
    function pendingUnknownEnforceModbusNetworkAndNcrSensorTypes() {
        if (activeTutorialMode !== EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) return;
        var puNcr = 11;
        var net = {
            id: "n1",
            name: "COM1 Modbus RTU",
            protocol: "Modbus/RTU",
            interface: "COM1",
            commParams: "9600,8,N,1",
            siteId: "st1"
        };
        state.sensorNetworks = (state.sensorNetworks || []).filter(function (n) {
            return n.siteId === "st1";
        });
        if (!state.sensorNetworks.length) {
            state.sensorNetworks = [net];
        } else {
            state.sensorNetworks[0] = net;
        }
        ensureNetworkFields(state.sensorNetworks[0]);
        state.vessels.forEach(function (v) {
            if (!v || v.siteId !== "st1") return;
            v.sensorNetworkId = "n1";
            v.sensorTypeId = puNcr;
        });
    }

    /** Pending / Unknown tutorial — single site, host IP matches workstation (no Database Read Only). */
    function applyPendingUnknownTutorialScenario() {
        state.sites = state.sites.filter(function (s) {
            return s.id === "st1";
        });
        if (!state.sites.length) {
            state.sites = [
                {
                    id: "st1",
                    name: "Current Location",
                    workstationSiteId: 1,
                    serviceHostIp: SIM_WORKSTATION_IPV4,
                    serviceHostPort: "8093",
                    companyName: "Acme Manufacturing",
                    streetAddress: "100 Industrial Pkwy",
                    streetAddress2: "",
                    city: "Minneapolis",
                    state: "MN",
                    zip: "55401",
                    country: "USA",
                    distanceUnitsId: "1"
                }
            ];
        }
        ensureSiteFields(state.sites[0]);
        state.sites[0].workstationSiteId = 1;
        state.sites[0].serviceHostIp = SIM_WORKSTATION_IPV4;
        state.currentWorkstationSiteId = "st1";
        state.sensorNetworks = (state.sensorNetworks || []).filter(function (n) {
            return !n.siteId || n.siteId === "st1";
        });
        state.groups = (state.groups || []).filter(function (g) {
            return !g.siteId || g.siteId === "st1";
        });
        state.schedules = (state.schedules || []).filter(function (s) {
            return !s.siteId || s.siteId === "st1";
        });
        seedWorkspaceMainPlant();
        pendingUnknownEnforceModbusNetworkAndNcrSensorTypes();
        var puNcr = 11;
        if (state.vessels[0]) {
            state.vessels[0].name = "NCR-80 Unknown";
            state.vessels[0].status = statusStringForDashboard(puNcr, 55);
            state.vessels[0].tutorialPuGlitch = true;
            delete state.vessels[0].tutorialPuStuck;
        }
        if (state.vessels[1]) {
            state.vessels[1].name = "NCR-80 Pending";
            state.vessels[1].status = statusStringForDashboard(puNcr, 90);
            state.vessels[1].tutorialPuStuck = true;
            delete state.vessels[1].tutorialPuGlitch;
        }
        state.vessels.forEach(function (v) {
            if (!v || v.siteId !== "st1") return;
            if (v.id === "v1" || v.id === "v2") return;
            v.status = statusStringForDashboard(puNcr, 0);
        });
        state._puGlitchDemoDone = false;
        state._puPostRestartResolved = false;
        state._puAwaitingPostRestartReadOnlyAck = false;
        delete state._puTutorialMeasureAllStarted;
    }

    function applyPendingUnknownPostRestartState(opts) {
        opts = opts || {};
        /* Tutorial recovery state after Services + client restart: connection restored and manual re-measure verifies both silos. */
        state.ebobServicesRunning = true;
        state.ebobSchedulerRunning = true;
        state.vesselsReadOnly = false;
        state._puPostRestartResolved = true;
        pendingUnknownEnforceModbusNetworkAndNcrSensorTypes();
        /* After services restart, all workstation-site silos read Unknown until measured (session restore can leave Ready/Pending). */
        var ws = state.currentWorkstationSiteId || "st1";
        state.vessels.forEach(function (v) {
            if (!v) return;
            if (v.siteId != null && v.siteId !== ws) return;
            delete v.tutorialPuGlitch;
            delete v.tutorialPuStuck;
            var sid = parseInt(v.sensorTypeId, 10) || 11;
            v.status = statusStringForDashboard(sid, 55);
            if (v.id === "v1" || v.id === "v2") {
                v.name = "NCR-80 Unknown";
            }
        });
        saveState();
        if (!opts.skipRefresh) {
            refreshUI();
        }
    }

    function restoreTutorialScenarioSessionState() {
        if (
            activeTutorialMode !== EBOB_TUTORIAL_MODE_DB_READ_ONLY &&
            activeTutorialMode !== EBOB_TUTORIAL_MODE_PENDING_UNKNOWN
        ) {
            return false;
        }
        try {
            var raw = sessionStorage.getItem(EBOB_TUTORIAL_STATE_KEY);
            if (!raw) return false;
            var parsed = JSON.parse(raw);
            if (!parsed || !parsed.stateSnapshot) return false;
            state = Object.assign(state, parsed.stateSnapshot);
            if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
                state._puGlitchDemoDone = false;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    function persistTutorialScenarioSessionState() {
        if (
            activeTutorialMode !== EBOB_TUTORIAL_MODE_DB_READ_ONLY &&
            activeTutorialMode !== EBOB_TUTORIAL_MODE_PENDING_UNKNOWN
        ) {
            return;
        }
        try {
            sessionStorage.setItem(
                EBOB_TUTORIAL_STATE_KEY,
                JSON.stringify({
                    mode: activeTutorialMode,
                    stateSnapshot: state
                })
            );
        } catch (e) {
            /* private mode */
        }
    }

    function loadState() {
        try {
            localStorage.removeItem(STORAGE_KEY);
        } catch (e) {
            /* private mode / blocked storage */
        }
        var wasImportRestart = false;
        try {
            wasImportRestart = sessionStorage.getItem(SIM_IMPORT_RESTART_FLAG) === "1";
        } catch (e2) {
            /* ignore */
        }
        activeTutorialMode = readTutorialMode();
        if (activeTutorialMode === EBOB_TUTORIAL_MODE_FREE) {
            activeTutorialMode = "";
            clearTutorialScenarioState();
        }
        resetStateToFactoryDefaults();
        if (activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY) {
            applyDatabaseReadOnlyTutorialScenario();
        } else if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
            applyPendingUnknownTutorialScenario();
        } else {
            clearTutorialScenarioState();
        }
        applyPendingImportRestartState();
        if (wasImportRestart) {
            if (state.ebobSchedulerRunning == null) {
                state.ebobSchedulerRunning = !!state.ebobServicesRunning;
            }
            persistEbobSvcSession();
        } else {
            /* Full refresh: no persisted service history — DB not read-only from last session. */
            try {
                sessionStorage.removeItem(EBOB_SVC_SESSION_KEY);
            } catch (e3) {
                /* private mode */
            }
        }
        syncReadOnlyLatchOnStartup();
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

    /** Clear frmLogin fields — user must enter credentials (matches eBob; no pre-filled passwords). */
    function clearLoginFormFields() {
        var u = document.getElementById("uid");
        var p = document.getElementById("pw");
        if (u) u.value = "";
        if (p) p.value = "";
    }

    function showLoginBackdrop() {
        clearLoginFormFields();
        var bdLoginEl = document.getElementById("backdropLogin");
        if (bdLoginEl) {
            bdLoginEl.classList.add("show");
            bdLoginEl.setAttribute("aria-hidden", "false");
        }
        setTimeout(function () {
            var u = document.getElementById("uid");
            if (u) u.focus();
        }, 0);
    }

    /**
     * frmMain.Login_Click — AutoLoginFlag = 1 → skip frmLogin; else frmLogin.ShowDialog (typed User ID + password).
     * On success: LoggedInMenuDefaults + SetupForm (refresh dashboard).
     */
    function performLoginFromMenu() {
        ensureSystemSettingsDefaults();
        var auto = state.systemSettings.autoLogin === true;
        if (auto) {
            var adm = state.users.filter(function (u) {
                return String(u.userId).toLowerCase() === "admin";
            })[0];
            if (adm) {
                state.currentUser = adm.name;
                adm.lastLogon = new Date().toLocaleString();
            } else {
                state.currentUser = null;
            }
            syncMenuStripForSession();
            refreshUI({ staggerVesselReveal: true });
            toast("Logged in (auto login).");
            return;
        }
        closeMenus();
        showLoginBackdrop();
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

    /**
     * Active alarm label for the vessel card — uses this vessel's enabled thresholds (BLL AlarmCondition semantics).
     * High / Pre-High: % full at or above threshold. Low / Pre-Low: % full at or below threshold.
     */
    function computeAlarm(v) {
        ensureVesselSetupDefaults(v);
        var pct = Number(v.pctFull);
        if (isNaN(pct)) pct = 0;
        pct = Math.max(0, Math.min(100, pct));
        var th;
        th = parseAlarmThresholdPct(v.alarmHighPct);
        if (v.alarmHighEnabled && th != null && pct >= th) return "High Alarm";
        th = parseAlarmThresholdPct(v.alarmPreHighPct);
        if (v.alarmPreHighEnabled && th != null && pct >= th) return "Pre-High Alarm";
        th = parseAlarmThresholdPct(v.alarmLowPct);
        if (v.alarmLowEnabled && th != null && pct <= th) return "Low Alarm";
        th = parseAlarmThresholdPct(v.alarmPreLowPct);
        if (v.alarmPreLowEnabled && th != null && pct <= th) return "Pre-Low Alarm";
        return "";
    }

    /**
     * Vessel.Designer.vb: tmrAlarm.Interval = 600. Vessel.vb tmrAlarm_Tick toggles lblAlarm.ForeColor
     * between Color.Black and Color.Red. Applied via inline color so no stylesheet / OS motion settings
     * can block the strobe in the browser.
     */
    var __ebobAlarmStrobePhaseBlack = true;

    function applyEbobAlarmStrobeColorToAllActive(c) {
        var nodes = document.querySelectorAll(".lbl-alarm.lbl-alarm--active");
        for (var i = 0; i < nodes.length; i++) {
            nodes[i].style.setProperty("color", c, "important");
        }
    }

    /** Re-apply current strobe phase after vessel cards are re-rendered (inline color survives stylesheet). */
    function syncEbobVesselAlarmStrobeInline() {
        var c = __ebobAlarmStrobePhaseBlack ? "#000000" : "#cc0000";
        applyEbobAlarmStrobeColorToAllActive(c);
    }

    function initVesselAlarmStrobeTimer() {
        if (typeof document === "undefined" || document.__ebobVesselAlarmStrobeInit) return;
        document.__ebobVesselAlarmStrobeInit = true;
        function tickVesselAlarmStrobe() {
            __ebobAlarmStrobePhaseBlack = !__ebobAlarmStrobePhaseBlack;
            var c = __ebobAlarmStrobePhaseBlack ? "#000000" : "#cc0000";
            applyEbobAlarmStrobeColorToAllActive(c);
        }
        document.__ebobVesselAlarmStrobeTimerId = setInterval(tickVesselAlarmStrobe, 600);
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
            toast(getReadOnlyMessage());
            return;
        }
        var sid = parseInt(v.sensorTypeId, 10);
        if (isNaN(sid)) sid = 2;
        /* SPL-100 / SPL-200 — Vessel.vb hides measure; readings come from scheduler / push only. */
        if (!opts.scheduled && (sid === 14 || sid === 15)) {
            if (!opts.silent) {
                toast("This sensor does not support manual measurement — use scheduled measurements.");
            }
            return;
        }
        /* Run before SmartBob/C-100MB so tutorial flags always win even if device type was clamped wrong. */
        if (
            activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN &&
            !isPendingUnknownPostRestartPhaseActive()
        ) {
            if (v.tutorialPuGlitch) {
                runPendingUnknownGlitchMeasurement(v, opts);
                return;
            }
            if (v.tutorialPuStuck) {
                runPendingUnknownStuckMeasurement(v, opts);
                return;
            }
        }
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

    /** Simulates “eBob services glitch”: Pending flashes, then back to Unknown (NCR-80). */
    function runPendingUnknownGlitchMeasurement(v, opts) {
        opts = opts || {};
        clearMeasureSimulation(v);
        var sid = parseInt(v.sensorTypeId, 10) || 11;
        v.status = statusStringForDashboard(sid, 90);
        saveState();
        renderGrid();
        v._measureTimers = [
            setTimeout(function () {
                v.status = statusStringForDashboard(sid, 55);
                state._puGlitchDemoDone = true;
                saveState();
                renderGrid();
                v._measureTimers = null;
                if (typeof opts.onDone === "function") opts.onDone();
            }, 220)
        ];
    }

    /** Pending forever — measurement never resolves to Ready. */
    function runPendingUnknownStuckMeasurement(v, opts) {
        opts = opts || {};
        clearMeasureSimulation(v);
        var sid = parseInt(v.sensorTypeId, 10) || 11;
        v.status = statusStringForDashboard(sid, 90);
        saveState();
        renderGrid();
        v._measureTimers = [
            setTimeout(function () {
                v.status = statusStringForDashboard(sid, 90);
                v._measureTimers = null;
                saveState();
                renderGrid();
                if (typeof opts.onDone === "function") opts.onDone();
            }, 400)
        ];
    }

    function isPendingUnknownPostRestartPhaseActive() {
        if (activeTutorialMode !== EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) return false;
        if (state._puPostRestartResolved) return true;
        if (!pendingUnknownTutorial || !pendingUnknownTutorial.active || pendingUnknownTutorial.completed) {
            return false;
        }
        /* Step 18+ — post-restart verification (first silo measure + Measure All). */
        return pendingUnknownTutorial.stepIndex >= 17;
    }

    /** Pending/Unknown tutorial: after Measure All, every measurable vessel on the workstation site should be Ready. */
    function pendingUnknownAllMeasurableSiteVesselsReady() {
        var siteId = state.currentWorkstationSiteId;
        var any = false;
        for (var i = 0; i < state.vessels.length; i++) {
            var v = state.vessels[i];
            if (!v || (v.siteId != null && v.siteId !== siteId)) continue;
            var sid = parseInt(v.sensorTypeId, 10);
            if (sid === 14 || sid === 15) continue;
            any = true;
            if (String(v.status || "").trim().toLowerCase() !== "ready") return false;
        }
        return any;
    }

    function measureAll() {
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to measure.");
            return;
        }
        if (state.vesselsReadOnly || !state.ebobServicesRunning) {
            toast(getReadOnlyMessage());
            return;
        }
        var list = state.vessels.filter(function (v) {
            var st = parseInt(v.sensorTypeId, 10);
            return st !== 14 && st !== 15;
        });
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
        var alarmRaw = computeAlarm(v);
        var alarm = alarmRaw || "\u00A0";
        var alarmActiveClass = alarmRaw ? " lbl-alarm--active" : "";
        var headChecked = v.headroom ? " checked" : "";
        var m = vesselDisplayMetrics(v);
        var row = v.headroom ? m.headroom : m.product;
        var disp = formatVolumeWeightPair(v, row.volumeCuFt, row.weightLb);
        var siloVars = siloFillCssVarsFromColorName(v.fillColor || "Dark Red");
        var sidCard = parseInt(v.sensorTypeId, 10);
        if (isNaN(sidCard)) sidCard = 2;
        var isSpl = sidCard === 14 || sidCard === 15;
        var addrRaw = v.sensorAddress != null ? String(v.sensorAddress).trim() : "";
        var contentsDisp =
            v.contents != null && String(v.contents).trim() !== ""
                ? String(v.contents).trim()
                : v.product != null && String(v.product).trim() !== ""
                  ? String(v.product).trim()
                  : "";
        if (!contentsDisp) contentsDisp = "\u00A0";
        var nameDisp =
            v.name != null && String(v.name).trim() !== "" ? String(v.name).trim() : "\u00A0";
        var siloTitle = addrRaw !== "" ? "Address " + addrRaw : "";
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
            '<div class="vessel-contents-line" title="' +
            escapeHtml(contentsDisp) +
            '">' +
            escapeHtml(contentsDisp) +
            "</div>" +
            "</div>" +
            "</div>" +
            '<div class="vessel-inner">' +
            '<div class="vessel-col-left">' +
            '<div class="vessel-graphic-col">' +
            '<div class="vessel-name-over-graphic" title="' +
            escapeHtml(nameDisp) +
            '">' +
            escapeHtml(nameDisp) +
            "</div>" +
            '<div class="silo" aria-hidden="true" style="' +
            escapeHtml(siloVars) +
            '" title="' +
            escapeHtml(siloTitle ? siloTitle + " — View vessel details" : "View vessel details") +
            '"' +
            ">" +
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
            escapeHtml(disp.volStr) +
            "</span>" +
            "<span>Weight:</span><span>" +
            escapeHtml(disp.wtStr) +
            "</span>" +
            "<span>Status:</span><span class=\"vessel-status-val" +
            statusCls +
            "\"><span class=\"vessel-status-text\">" +
            escapeHtml(statusText) +
            "</span></span>" +
            "</div>" +
            '<label class="chk-row"><input type="checkbox" data-vessel-action="headroom"' +
            headChecked +
            '> Headroom Display</label>' +
            '<div class="lbl-alarm' +
            alarmActiveClass +
            '">' +
            escapeHtml(alarm) +
            "</div>" +
            (isSpl
                ? ""
                : '<div class="vessel-measure-row">' +
                  '<button type="button" class="btn-measure"' +
                  measureDis +
                  ' data-vessel-action="measure">Measure</button>' +
                  "</div>") +
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
                var newPage = parseInt(this.dataset.tab, 10);
                if (newPage === state.currentPage) return;
                state.currentPage = newPage;
                saveState();
                renderTabs();
                runDashboardPageTabSwitch();
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
                syncEbobVesselAlarmStrobeInline();
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
        syncEbobVesselAlarmStrobeInline();
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
                "Lost connection to database — read only mode. Close Binventory and restart the workstation after starting the eBob Engine service."
            );
        } else if (state.ebobServicesRunning && !was) {
            if (state.vesselsReadOnly) {
                toast(
                    "eBob Engine is running. Close and restart Binventory to exit read-only mode."
                );
            } else {
                toast("Database connection restored — full access.");
            }
        }
        persistEbobSvcSession();
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

    /**
     * Dashboard Page 1 / Page 2 — staggered tiles with Loading… then bind (no full-screen wait; faster than site switch).
     */
    function runDashboardPageTabSwitch() {
        renderGrid({
            staggerVesselReveal: true,
            staggerDelayMs: 44,
            staggerLoadingDwellMs: 560
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
            "backdropDeviceMgr",
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
        var svcBd = document.getElementById("backdropServicesMsc");
        if (svcBd) svcBd.classList.remove("backdrop-services--minimized");
        var tbs = document.getElementById("taskbarBtnServices");
        if (tbs) {
            tbs.hidden = true;
            tbs.classList.remove("win-taskbar-services--active");
        }
        var dmBd = document.getElementById("backdropDeviceMgr");
        if (dmBd) dmBd.classList.remove("backdrop-device-mgr--minimized");
        var tbdm = document.getElementById("taskbarBtnDeviceMgr");
        if (tbdm) {
            tbdm.hidden = true;
            tbdm.classList.remove("win-taskbar-device-mgr--active");
        }
        var bdAppExit = document.getElementById("backdropApp");
        if (bdAppExit) bdAppExit.classList.remove("backdrop-app--terminal-minimized");
        var tbt = document.getElementById("taskbarBtnTerminal");
        if (tbt) {
            tbt.hidden = true;
            tbt.classList.remove("win-taskbar-terminal--active");
        }
    }

    function exitToSimDesktop() {
        persistTutorialScenarioSessionState();
        if (activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY && isGuidedTutorialLockActive()) {
            var gtExit = guidedGetActiveTutorialState();
            if (gtExit) {
                var stExit = gtExit.steps[gtExit.stepIndex];
                if (stExit && stExit.dbRoRequireRecordedExit) {
                    state._dbRoTutorialClosedAppForStep16 = true;
                }
            }
        }
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
                var restoredScenarioState = restoreTutorialScenarioSessionState();
                if (!restoredScenarioState) {
                    resetStateToFactoryDefaults();
                    if (activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY) {
                        applyDatabaseReadOnlyTutorialScenario();
                    } else if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
                        applyPendingUnknownTutorialScenario();
                    }
                } else if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
                    pendingUnknownEnforceModbusNetworkAndNcrSensorTypes();
                }
                applyEbobSvcSession();
                syncReadOnlyLatchOnStartup();
                /* Pending/Unknown: relaunch after Close + Relaunch sets full grid (Unknown / Ready path); mid-tutorial desktop exit before that uses the flag. */
                if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
                    if (state._puAwaitingPostRestartReadOnlyAck) {
                        state._puAwaitingPostRestartReadOnlyAck = false;
                        applyPendingUnknownPostRestartState({ skipRefresh: true });
                    } else if (restoredScenarioState && state._puPostRestartResolved) {
                        applyPendingUnknownPostRestartState({ skipRefresh: true });
                    }
                }
                var shell = document.getElementById("appShell");
                var tbEbob = document.getElementById("taskbarBtnEbob");
                if (shell) shell.classList.remove("app-shell--minimized");
                if (tbEbob) tbEbob.classList.remove("win-taskbar-ebob--minimized");
                applyAutoLoginSessionIfEnabled();
                refreshUI({ staggerVesselReveal: true });
                if (state.vesselsReadOnly || !state.ebobServicesRunning) {
                    toast(getReadOnlyMessage());
                } else {
                    toast("Connected to database — eBob Workstation is ready.");
                }
                startDbReadOnlyTutorialIfNeeded();
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
    var appCapMin = document.getElementById("appCapMin");
    var appCapMax = document.getElementById("appCapMax");
    var taskbarBtnTerminal = document.getElementById("taskbarBtnTerminal");

    /** Second layer (z-index) — frmScheduleMaintenance shown over frmEmailReports (WinForms ShowDialog stack). */
    var bdAppStack = document.getElementById("backdropAppStack");
    var appModalShellStack = document.getElementById("appModalShellStack");
    var appModalTitleStack = document.getElementById("appModalTitleStack");
    var appModalBodyStack = document.getElementById("appModalBodyStack");
    var appModalFooterStack = document.getElementById("appModalFooterStack");

    /** Third layer — Measurement Schedule Setup / Assign Groups over schedule list (parent stays visible). */
    var bdAppStack2 = document.getElementById("backdropAppStack2");
    var appModalShellStack2 = document.getElementById("appModalShellStack2");
    /** When Assign Vessels (#backdropAppStack2) closes, return to Vessel Group Setup (caption ✕ or Close). */
    var vesselGroupAssignCallback = null;
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
        shell.classList.remove("modal-user-setup");
        shell.classList.remove("modal-contact-maint");
        shell.classList.remove("modal-contact-setup");
        shell.classList.remove("modal-group-setup");
        shell.classList.remove("modal-email-setup");
        shell.classList.remove("modal-email-reports");
        shell.classList.remove("modal-site-setup");
        shell.classList.remove("modal-win-toolwindow");
        shell.classList.remove("modal-vessel-details");
    }

    bdApp.addEventListener("click", function (e) {
        if (e.target.closest("[data-mock-report-print]")) {
            e.preventDefault();
            printMockReportWindow();
            return;
        }
        if (e.target.closest("#appCapMin")) {
            e.preventDefault();
            minimizeTerminalApp();
            return;
        }
        if (e.target.closest("#appCapMax")) {
            e.preventDefault();
            if (appModalShell && appModalShell.classList.contains("modal-command-prompt")) {
                appModalShell.classList.toggle("modal-command-prompt--max");
            }
            return;
        }
        if (e.target.closest("[data-close-app]")) {
            var vsOnStack =
                vmSetupEditingId &&
                bdAppStack &&
                bdAppStack.classList.contains("show") &&
                appModalShellStack.classList.contains("modal-vessel-setup");
            var vsOnBase = appModalShell.classList.contains("modal-vessel-setup") && vmSetupEditingId;
            if (vsOnStack || vsOnBase) {
                closeVesselSetupDiscardChanges();
            } else {
                closeAppModal();
            }
        }
    });

    document.addEventListener(
        "keydown",
        function (e) {
            if (e.key !== "Escape") return;
            if (!bdApp || !bdApp.classList.contains("show")) return;
            if (!appModalShell || !appModalShell.classList.contains("modal-command-prompt")) return;
            var tag = e.target && e.target.tagName;
            if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
            e.preventDefault();
            e.stopPropagation();
            minimizeTerminalApp();
        },
        true
    );

    if (taskbarBtnTerminal) {
        taskbarBtnTerminal.addEventListener("click", function () {
            if (!bdApp) return;
            if (bdApp.classList.contains("show")) return;
            restoreTerminalApp();
        });
    }

    window.__ebobRestoreTerminalApp = restoreTerminalApp;

    if (bdAppStack) {
        bdAppStack.addEventListener("click", function (e) {
            if (e.target.closest("[data-close-app-stack]")) {
                if (appModalShellStack.classList.contains("modal-vessel-setup") && vmSetupEditingId) {
                    closeVesselSetupDiscardChanges();
                } else {
                    closeStackedAppModal();
                }
            }
        });
    }

    function finishVesselGroupAssignAndReturn() {
        var cb = vesselGroupAssignCallback;
        vesselGroupAssignCallback = null;
        closeStackedAppModal2();
        if (typeof cb === "function") cb();
    }

    if (bdAppStack2) {
        bdAppStack2.addEventListener("click", function (e) {
            if (!e.target.closest("[data-close-app-stack2]")) return;
            if (
                appModalShellStack2.classList.contains("modal-vessel-group-assign") &&
                vesselGroupAssignCallback != null
            ) {
                finishVesselGroupAssignAndReturn();
                return;
            }
            closeStackedAppModal2();
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

    function syncTerminalTaskbarButton(visible) {
        if (!taskbarBtnTerminal) return;
        if (visible) {
            taskbarBtnTerminal.hidden = false;
            taskbarBtnTerminal.classList.add("win-taskbar-terminal--active");
        } else {
            taskbarBtnTerminal.hidden = true;
            taskbarBtnTerminal.classList.remove("win-taskbar-terminal--active");
        }
    }

    function minimizeTerminalApp() {
        if (!bdApp || !appModalShell || !appModalShell.classList.contains("modal-command-prompt")) return;
        bdApp.classList.remove("show");
        bdApp.classList.add("backdrop-app--terminal-minimized");
        syncTerminalTaskbarButton(true);
    }

    function restoreTerminalApp() {
        if (!bdApp || !appModalShell || !appModalShell.classList.contains("modal-command-prompt")) return;
        bdApp.classList.remove("backdrop-app--terminal-minimized");
        bdApp.classList.add("show");
        syncTerminalTaskbarButton(true);
        setTimeout(function () {
            var inp = document.getElementById("cmdInput");
            if (inp && inp.focus) inp.focus();
        }, 0);
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
        bdApp.classList.remove("backdrop-app--terminal-minimized");
        syncTerminalTaskbarButton(false);
        appModalShell.classList.remove("modal-command-prompt--max");
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
        appModalShell.classList.remove("modal-user-setup");
        appModalShell.classList.remove("modal-contact-maint");
        appModalShell.classList.remove("modal-contact-setup");
        appModalShell.classList.remove("modal-group-setup");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-win-toolwindow");
        appModalShell.classList.remove("modal-command-prompt");
        appModalShell.classList.remove("modal-vessel-details");
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
        appModalShell.classList.remove("modal-user-setup");
        appModalShell.classList.remove("modal-contact-maint");
        appModalShell.classList.remove("modal-contact-setup");
        appModalShell.classList.remove("modal-group-setup");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-win-toolwindow");
        appModalShell.classList.remove("modal-command-prompt");
        appModalShell.classList.remove("modal-command-prompt--max");
        appModalShell.classList.remove("modal-vessel-details");
        if (modalClass) {
            String(modalClass)
                .split(/\s+/)
                .forEach(function (c) {
                    if (c) appModalShell.classList.add(c);
                });
        }
        bdApp.classList.remove("backdrop-app--terminal-minimized");
        if (appModalShell.classList.contains("modal-command-prompt")) {
            syncTerminalTaskbarButton(true);
        } else {
            syncTerminalTaskbarButton(false);
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

    /** frmContactSetup / ContactRecord — FirstName, LastName, EmailAddress, JobTitle; display ContactName. */
    function ensureContactRecord(c) {
        if (!c || typeof c !== "object") return c;
        if (c.firstName == null) c.firstName = "";
        if (c.lastName == null) c.lastName = "";
        if (c.jobTitle == null) c.jobTitle = "";
        if (c.email == null) c.email = "";
        if (c.emailAddress == null) c.emailAddress = c.email;
        else c.email = String(c.emailAddress || c.email || "");
        c.emailAddress = c.email;
        var fn0 = String(c.firstName || "").trim();
        var ln0 = String(c.lastName || "").trim();
        if (!fn0 && !ln0 && c.name) {
            var parts = String(c.name || "").trim().split(/\s+/);
            if (parts.length) {
                c.firstName = parts[0];
                c.lastName = parts.length > 1 ? parts.slice(1).join(" ") : "";
            }
        }
        var fn = String(c.firstName || "").trim();
        var ln = String(c.lastName || "").trim();
        if (fn || ln) {
            c.name = (fn + " " + ln).trim();
        }
        if (!c.name) c.name = "Contact";
        return c;
    }

    function contactNameParts(c) {
        if (!c) return { first: "", last: "" };
        ensureContactRecord(c);
        return { first: String(c.firstName || "").trim(), last: String(c.lastName || "").trim() };
    }

    function ensureVesselContactIds(v) {
        if (!v) return;
        if (!Array.isArray(v.vesselContactIds)) v.vesselContactIds = [];
    }

    function removeContactIdFromAllVessels(contactId) {
        state.vessels.forEach(function (v) {
            ensureVesselContactIds(v);
            v.vesselContactIds = v.vesselContactIds.filter(function (id) {
                return id !== contactId;
            });
        });
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
        var tbody = vesselSetupField("vs_email_contacts_tbody");
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
            toast("Vessel deleted.");
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

    /** frmVesselSetup.cboContents — suggestions only; user can type any value (MaxLength 25 in BLL). */
    function buildVmSetupContentsDatalistOptions(current) {
        var cur = String(current || "").trim();
        var seen = {};
        var parts = PRODUCTS.map(function (p) {
            seen[p] = true;
            return '<option value="' + escapeHtml(p) + '"></option>';
        });
        if (cur && !seen[cur]) {
            parts.unshift('<option value="' + escapeHtml(cur) + '"></option>');
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

    /** Display like eBob: "Name (Protocol on COM2)" or IP for remote. */
    function formatNetworkSelectLabel(n) {
        if (!n) return "";
        if (!n.name) n.name = "Sensor Network";
        if (!n.protocol) n.protocol = "Protocol A";
        if (!n.interface) n.interface = "COM2";
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
        var fillSel = v.fillColor || "Dark Red";
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
            '<fieldset class="vs-group vs-group-ebob-contents">' +
            '<legend class="vs-group-legend">Vessel Contents</legend>' +
            '<div class="vs-contents-abs-host">' +
            '<label class="vs-contents-abs-lbl" for="vs_contents">Contents / Product:</label>' +
            '<div class="vs-contents-abs-combo">' +
            '<div class="vs-combo vs-combo-contents">' +
            '<input type="text" id="vs_contents" class="vs-combo-input" list="vs_contents_list" maxlength="' +
            EBOB_VESSEL_CONTENTS_MAX +
            '" spellcheck="false" autocomplete="off" value="' +
            escapeHtml(String(contents != null ? contents : "")) +
            '">' +
            "</div>" +
            '<datalist id="vs_contents_list">' +
            buildVmSetupContentsDatalistOptions(contents) +
            "</datalist></div></div>" +
            "</fieldset>" +
            '<fieldset class="vs-group vs-group-ebob-contents">' +
            '<legend class="vs-group-legend">Density / Specific Gravity of Product</legend>' +
            '<div class="vs-den-controls-ebob">' +
            '<div class="vs-den-abs-rad vs-den-abs-rad-pd">' +
            '<input type="radio" name="vs_den_mode" id="vs_den_density" value="density"' +
            (densityChecked ? " checked" : "") +
            '>' +
            '<label for="vs_den_density">Use Product Density:</label>' +
            "</div>" +
            '<input type="text" id="vs_density_val" class="vs-input vs-den-abs-txt vs-den-abs-density" value="' +
            escapeHtml(String(v.productDensity || "")) +
            '">' +
            '<div id="vs_density_units_row" class="vs-den-abs-units-cluster">' +
            '<label class="vs-den-abs-lbl-units" for="vs_density_units" id="vs_density_units_lbl">Density Units:</label>' +
            '<select id="vs_density_units" class="vs-select vs-den-abs-cbo-units">' +
            buildSelectOptionsStringList(DENSITY_UNITS_OPTIONS, v.densityUnits) +
            "</select></div>" +
            '<div class="vs-den-abs-rad vs-den-abs-rad-sg">' +
            '<input type="radio" name="vs_den_mode" id="vs_den_sg" value="sg"' +
            (sgChecked ? " checked" : "") +
            '>' +
            '<label for="vs_den_sg">Use Specific Gravity:</label>' +
            "</div>" +
            '<input type="text" id="vs_sg_val" class="vs-input vs-den-abs-txt vs-den-abs-sg" value="' +
            escapeHtml(String(v.specificGravity != null ? v.specificGravity : "")) +
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
            '<select id="vs_sensor_distvar" class="vs-select vs-select-distvar">' +
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
        var root = vesselSetupDomRoot();
        var shell = root && root.querySelector(".vs-shell");
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

        var nameEl = vesselSetupField("vs_name");
        var titleEl = vesselSetupField("vs_lbl_title");
        if (nameEl && titleEl) {
            nameEl.addEventListener("input", function () {
                titleEl.textContent = "Vessel Setup - " + nameEl.value;
            });
        }

        var colorEl = vesselSetupField("vs_color");
        var prevEl = vesselSetupField("vs_color_preview");
        if (colorEl && prevEl) {
            colorEl.addEventListener("change", function () {
                prevEl.style.background = fillColorPreviewCss(colorEl.value);
            });
        }

        var typeEl = vesselSetupField("vs_vessel_type");
        var asstEl = vesselSetupField("vs_assistance");
        var previewEl = vesselSetupField("vs_type_preview");
        var shapeFieldsSlot = vesselSetupField("vs_shape_fields_slot");
        var partChk = vesselSetupField("vs_partition_chk");
        var partScale = vesselSetupField("vs_partition_scale");
        var partPct = document.querySelector(".vs-part-pct");
        var partLbl = document.querySelector(".vs-lbl-part");

        function syncTypeShapeDiagram() {
            if (!previewEl || !typeEl) return;
            var tid = parseInt(typeEl.value, 10) || 1;
            if (tid === 15) {
                previewEl.innerHTML = "";
                return;
            }
            var hEl = vesselSetupField("vs_sp_0");
            var wEl = vesselSetupField("vs_sp_1");
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
            var pnl = vesselSetupField("vs_pnl_split");
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
                        var u = vesselSetupField("vs_lbl_dist_val");
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
                var tbSel = vesselSetupField("vs_custom_tbody");
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
                    var tb = vesselSetupField("vs_custom_tbody");
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
                    var tb2 = vesselSetupField("vs_custom_tbody");
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
                    var tbImp = vesselSetupField("vs_custom_tbody");
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
                var th = vesselSetupField("vs_custom_th_out");
                if (th) {
                    th.textContent = customOutputColumnHeader(idx, da, va, wa);
                }
            });
        }

        var distRow = vesselSetupField("vs_dist_row");
        if (distRow) {
            distRow.addEventListener("mouseenter", function () {
                var u = vesselSetupField("vs_lbl_dist_val");
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

        var emailEn = vesselSetupField("vs_email_en");
        var grpEmail = shell.querySelector(".vs-email-grid");
        var selBtn = vesselSetupField("vs_sel_contacts");
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

        function syncDensitySgControls() {
            var sgOn = vesselSetupField("vs_den_sg") && vesselSetupField("vs_den_sg").checked;
            var dVal = vesselSetupField("vs_density_val");
            var dUnits = vesselSetupField("vs_density_units");
            var sgVal = vesselSetupField("vs_sg_val");
            if (dVal) dVal.disabled = !!sgOn;
            if (dUnits) dUnits.disabled = !!sgOn;
            if (sgVal) sgVal.disabled = !sgOn;
            var lblDu = vesselSetupField("vs_density_units_lbl");
            if (lblDu) lblDu.classList.toggle("vs-disabled", !!sgOn);
            var rowDu = vesselSetupField("vs_density_units_row");
            if (rowDu) rowDu.classList.toggle("vs-disabled", !!sgOn);
        }
        shell.addEventListener("change", function (e) {
            if (e.target && e.target.name === "vs_den_mode") syncDensitySgControls();
        });
        syncDensitySgControls();

        if (selBtn) {
            selBtn.addEventListener("click", function () {
                if (!emailEn || !emailEn.checked) return;
                openAssignContactsDialog();
            });
        }

        var sensorTypeEl = vesselSetupField("vs_sensor_type");
        var networkEl = vesselSetupField("vs_network");

        function syncSbMaxDrop() {
            var rowOn = vesselSetupField("vs_sb_en1") && vesselSetupField("vs_sb_en1").checked;
            var mdEn = vesselSetupField("vs_sb_en_maxdrop") && vesselSetupField("vs_sb_en_maxdrop").checked;
            var md = vesselSetupField("vs_sb_maxdrop");
            var em = vesselSetupField("vs_sb_en_maxdrop");
            if (em) em.disabled = !rowOn;
            if (md) md.disabled = !rowOn || !mdEn;
        }

        function syncSbRowEnabled() {
            var en = vesselSetupField("vs_sb_en1");
            var addr = vesselSetupField("vs_sb_addr1");
            var off = vesselSetupField("vs_sb_offset1");
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
            var gen = vesselSetupField("vs_sensor_pnl_generic");
            var sb = vesselSetupField("vs_sensor_pnl_smartbob");
            var vg = vesselSetupField("vs_sensor_pnl_vega");
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
            var rowDec = vesselSetupField("vs_row_decimal");
            var rowDv = vesselSetupField("vs_row_distvar");
            var rowOff = vesselSetupField("vs_row_sensor_offset");
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
            var en = vesselSetupField("vs_sensor_enabled");
            var addr = vesselSetupField("vs_sensor_address");
            var off = vesselSetupField("vs_sensor_offset");
            var on = en && en.checked;
            if (addr) addr.disabled = !on;
            if (off) off.disabled = !on;
        }

        var sensorEnChk = vesselSetupField("vs_sensor_enabled");
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

        var sbEn1 = vesselSetupField("vs_sb_en1");
        var sbEnMax = vesselSetupField("vs_sb_en_maxdrop");
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
                var enG = vesselSetupField("vs_vg_en" + (ix + 1));
                var aG = vesselSetupField("vs_vg_a" + (ix + 1));
                var oG = vesselSetupField("vs_vg_o" + (ix + 1));
                var dG = vesselSetupField("vs_vg_dv" + (ix + 1));
                if (enG) vv.vegaSensors[ix].enabled = enG.checked;
                if (aG) vv.vegaSensors[ix].address = aG.value;
                if (oG) vv.vegaSensors[ix].offset = oG.value;
                if (dG) vv.vegaSensors[ix].dv = dG.value;
            }
        }

        function rebuildVegaSensorTable() {
            var vv = findVessel(vmSetupEditingId);
            if (!vv) return;
            var vcEl = vesselSetupField("vs_vega_count");
            var tbody = vesselSetupField("vs_vega_tbody");
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
            var vcEl = vesselSetupField("vs_vega_count");
            var nR = vcEl ? Math.min(32, Math.max(1, parseInt(vcEl.value, 10) || 1)) : 1;
            var r;
            for (r = 1; r <= nR; r++) {
                (function (rowIdx) {
                    var enR = vesselSetupField("vs_vg_en" + rowIdx);
                    var aR = vesselSetupField("vs_vg_a" + rowIdx);
                    var oR = vesselSetupField("vs_vg_o" + rowIdx);
                    var dR = vesselSetupField("vs_vg_dv" + rowIdx);
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

        var vegaCountEl = vesselSetupField("vs_vega_count");
        if (vegaCountEl) {
            vegaCountEl.addEventListener("change", rebuildVegaSensorTable);
        }
        bindVegaRowEnableHandlers();
    }

    function closeVesselSetupToMaintenance() {
        dismissAssignContactsBackdrop();
        closeStackedAppModal();
        appModalShell.classList.remove("modal-vessel-setup");
        appModalShell.classList.remove("modal-vessel-details");
        appModalShell.classList.remove("modal-sn-networks");
        appModalShell.classList.remove("modal-system-setup");
        appModalShell.classList.remove("modal-site-maintenance");
        appModalShell.classList.remove("modal-site-setup");
        appModalShell.classList.remove("modal-report-preview");
        appModalShell.classList.remove("modal-report-crystal");
        appModalShell.classList.remove("modal-email-reports");
        appModalShell.classList.remove("modal-email-setup");
        appModalShell.classList.remove("modal-user-maintenance");
        appModalShell.classList.remove("modal-user-setup");
        appModalShell.classList.remove("modal-contact-maint");
        appModalShell.classList.remove("modal-contact-setup");
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
        openStackedAppModal(
            "Vessel Setup — Binventory Workstation",
            buildVesselSetupBody(v),
            '<button type="button" class="primary" id="vsSave">Save</button><button type="button" class="secondary" id="vsCancel">Cancel</button>',
            "modal-vessel-setup"
        );
        bindVesselSetupForm();

        document.getElementById("vsSave").addEventListener("click", function () {
            var vv = findVessel(vmSetupEditingId);
            if (!vv) return;
            var dispErr = validateVesselSetupDisplayFromDom();
            if (!dispErr.ok) {
                toast(dispErr.message);
                return;
            }
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
            vv.name = (vesselSetupField("vs_name") && vesselSetupField("vs_name").value.trim()) || vv.name;
            vv.fillColor = vesselSetupField("vs_color") ? vesselSetupField("vs_color").value : vv.fillColor;
            vv.volumeDisplayUnits = vesselSetupField("vs_vol_units")
                ? vesselSetupField("vs_vol_units").value
                : vv.volumeDisplayUnits;
            vv.weightDisplayUnits = vesselSetupField("vs_wt_units")
                ? vesselSetupField("vs_wt_units").value
                : vv.weightDisplayUnits;
            vv.defaultHeadroom = vesselSetupField("vs_headroom")
                ? vesselSetupField("vs_headroom").checked
                : false;
            vv.headroom = vv.defaultHeadroom;
            vv.vesselTypeId = parseInt(vesselSetupField("vs_vessel_type").value, 10) || 1;
            refreshVesselShapeParamsFromDom(vv);
            if (vesselSetupField("vs_partition_chk")) {
                vv.verticalSplitPartitioned = vesselSetupField("vs_partition_chk").checked;
            }
            if (vesselSetupField("vs_partition_scale")) {
                vv.partitionScale = vesselSetupField("vs_partition_scale").value;
            }
            if (!vesselTypeShowsVerticalSplitPartition(vv.vesselTypeId)) {
                vv.verticalSplitPartitioned = false;
                vv.partitionScale = "100";
            }
            var cSel = vesselSetupField("vs_contents");
            vv.contents = cSel ? getVsContentsTextFromDom() || vv.contents : vv.contents;
            vv.product = vv.contents;
            vv.densityMode = vesselSetupField("vs_den_sg") && vesselSetupField("vs_den_sg").checked ? "sg" : "density";
            var dDen = vesselSetupField("vs_density_val");
            vv.productDensity = dDen ? String(dDen.value).trim() : vv.productDensity;
            vv.densityUnits = vesselSetupField("vs_density_units")
                ? vesselSetupField("vs_density_units").value
                : vv.densityUnits;
            var sgEl = vesselSetupField("vs_sg_val");
            vv.specificGravity = sgEl ? String(sgEl.value).trim() : "";
            vv.heightFt = parseFloat(vesselSetupField("vs_h").value) || vv.heightFt;
            vv.volumeCuFt = parseInt(vesselSetupField("vs_vol").value, 10) || vv.volumeCuFt;
            vv.weightLb = parseInt(vesselSetupField("vs_w").value, 10) || vv.weightLb;
            vv.pctFull = Math.min(
                100,
                Math.max(0, parseInt(vesselSetupField("vs_pct").value, 10) || vv.pctFull)
            );
            vv.sensorNetworkId = vesselSetupField("vs_network")
                ? vesselSetupField("vs_network").value || vv.sensorNetworkId
                : vv.sensorNetworkId;
            vv.sensorTypeId = parseInt(vesselSetupField("vs_sensor_type").value, 10) || 1;
            var stSave = vv.sensorTypeId;
            var bSave = getSensorAddressBounds(stSave);
            if (stSave === 1 || stSave === 3) {
                if (vesselSetupField("vs_sb_en1")) {
                    vv.sensorEnabled = vesselSetupField("vs_sb_en1").checked;
                }
                if (vesselSetupField("vs_sb_addr1")) {
                    vv.sensorAddress = validateIntegerSensorAddress(
                        vesselSetupField("vs_sb_addr1").value,
                        bSave.min,
                        bSave.max
                    ).normalized;
                }
                if (vesselSetupField("vs_sb_offset1")) {
                    vv.sensorOffset = vesselSetupField("vs_sb_offset1").value;
                }
                if (vesselSetupField("vs_sb_en_maxdrop")) {
                    vv.sbEnableMaxDrop = vesselSetupField("vs_sb_en_maxdrop").checked;
                }
                if (vesselSetupField("vs_sb_maxdrop")) {
                    vv.sbMaxDrop = vesselSetupField("vs_sb_maxdrop").value;
                }
            } else if (stSave === 11 || stSave === 12) {
                ensureVegaSensorsArray(vv);
                var vcSg = vesselSetupField("vs_vega_count");
                var nSg = Math.min(32, Math.max(1, parseInt(vcSg && vcSg.value, 10) || 1));
                vv.vegaSensorCount = String(nSg);
                var vx;
                for (vx = 0; vx < 32; vx++) {
                    var enSg = vesselSetupField("vs_vg_en" + (vx + 1));
                    var aSg = vesselSetupField("vs_vg_a" + (vx + 1));
                    var oSg = vesselSetupField("vs_vg_o" + (vx + 1));
                    var dSg = vesselSetupField("vs_vg_dv" + (vx + 1));
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
                if (vesselSetupField("vs_sensor_enabled")) {
                    vv.sensorEnabled = vesselSetupField("vs_sensor_enabled").checked;
                }
                if (vesselSetupField("vs_sensor_address")) {
                    vv.sensorAddress = validateIntegerSensorAddress(
                        vesselSetupField("vs_sensor_address").value,
                        bSave.min,
                        bSave.max
                    ).normalized;
                }
                if (vesselSetupField("vs_sensor_offset")) {
                    vv.sensorOffset = vesselSetupField("vs_sensor_offset").value;
                }
            }
            if (vesselSetupField("vs_sensor_decimal")) {
                vv.sensorDecimalPlaces = vesselSetupField("vs_sensor_decimal").value;
            }
            if (vesselSetupField("vs_sensor_distvar")) {
                vv.sensorDistanceVariable = vesselSetupField("vs_sensor_distvar").value;
            }
            vv.alarmHighEnabled = vesselSetupField("vs_ah_high").checked;
            vv.alarmHighPct = vesselSetupField("vs_ah_high_pct").value;
            vv.alarmPreHighEnabled = vesselSetupField("vs_ah_prehigh").checked;
            vv.alarmPreHighPct = vesselSetupField("vs_ah_prehigh_pct").value;
            vv.alarmPreLowEnabled = vesselSetupField("vs_ah_prelow").checked;
            vv.alarmPreLowPct = vesselSetupField("vs_ah_prelow_pct").value;
            vv.alarmLowEnabled = vesselSetupField("vs_ah_low").checked;
            vv.alarmLowPct = vesselSetupField("vs_ah_low_pct").value;
            vv.emailNotificationsEnabled = vesselSetupField("vs_email_en").checked;
            vv.emailFlags = {
                high: vesselSetupField("vs_ef_high").checked,
                preHigh: vesselSetupField("vs_ef_prehigh").checked,
                preLow: vesselSetupField("vs_ef_prelow").checked,
                low: vesselSetupField("vs_ef_low").checked,
                vesselStatus: vesselSetupField("vs_ef_status").checked,
                error: vesselSetupField("vs_ef_err").checked
            };
            saveState();
            delete vesselSetupSnapshotById[vmSetupEditingId];
            closeVesselSetupToMaintenance();
            refreshUI();
            toast("Vessel saved.");
        });
        document.getElementById("vsCancel").addEventListener("click", function () {
            closeVesselSetupDiscardChanges();
        });
    }

    function openVesselMaintenance() {
        state.vessels.forEach(function (v, idx) {
            ensureVesselFields(v, idx);
        });
        resolveDuplicateSensorAddressesOnAllVessels();
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

    /** Mirrors giAccessLevel — accessLevel 3 = Read-Only; logged out → ReadOnly (Logoff_Click). */
    function currentUserAccessLevel() {
        if (!isSessionLoggedIn()) return "readonly";
        var u = state.users.filter(function (x) {
            return x.name === state.currentUser || String(x.userId) === state.currentUser;
        })[0];
        if (!u) return "readonly";
        ensureUserRecord(u);
        if (u.accessLevel === 3) return "readonly";
        return "admin";
    }

    /**
     * frmVesselGroup — Assign Vessels to Group (584×439). Mutates working.vesselIds; onClose returns to Group Setup.
     */
    function openVesselGroupAssignDialog(working, onClose) {
        vesselGroupAssignCallback = typeof onClose === "function" ? onClose : null;
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

        openStackedAppModal2(
            "Assign Vessels to Group - Binventory Workstation",
            html,
            "",
            "modal-vessel-group-assign modal-footer-hidden modal-win-toolwindow"
        );

        var vgaShell = document.querySelector("#appModalBodyStack2 .vga-shell") || document.querySelector(".vga-shell");
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
            finishVesselGroupAssignAndReturn();
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
                    closeStackedAppModal();
                    refreshGroupMaintenanceContent({});
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

        openStackedAppModal(
            "Vessel Group Setup - Binventory Workstation",
            html,
            "",
            "modal-group-setup modal-footer-hidden modal-win-toolwindow"
        );

        renderGsVesselRows();

        document.getElementById("gsAssign").addEventListener("click", function () {
            if (ro) return;
            working.name = document.getElementById("gsName").value.trim();
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
            closeStackedAppModal();
            if (wasNew) {
                refreshGroupMaintenanceContent({ selectGroupId: working.id });
            } else {
                refreshGroupMaintenanceContent(
                    restoreParent
                        ? Object.assign({}, restoreParent, {
                              selectGroupId: working.id
                          })
                        : { selectGroupId: working.id }
                );
            }
            toast("Group saved.");
        });
        document.getElementById("gsCancel").addEventListener("click", function () {
            closeStackedAppModal();
            refreshGroupMaintenanceContent(restoreParent || {});
        });
    }

    function buildGroupMaintenanceBodyHtml() {
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

        return (
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
            "</div></div></div>"
        );
    }

    function refreshGroupMaintenanceContent(opts) {
        opts = opts || {};
        if (!appModalShell.classList.contains("modal-group-maint")) {
            openGroupMaintenance(opts);
            return;
        }
        updateAppModalContent(
            "Group Maintenance - Binventory Workstation",
            buildGroupMaintenanceBodyHtml(),
            ""
        );
        bindGroupMaintenanceEvents(opts);
    }

    function bindGroupMaintenanceEvents(opts) {
        opts = opts || {};
        var ro = currentUserAccessLevel() === "readonly";
        var tbody = document.querySelector("#gmGroupTable tbody");
        var wrap = document.querySelector(".gm-dgv-wrap");
        var btnSel = document.getElementById("gmBtnSelect");
        var btnDel = document.getElementById("gmBtnDelete");

        function selectedGroupId() {
            var tr = tbody && tbody.querySelector("tr.gm-row.selected");
            return tr ? tr.getAttribute("data-group-id") : null;
        }

        function syncGmButtons() {
            var id = selectedGroupId();
            var on = !!id;
            if (btnSel) btnSel.disabled = !on;
            if (btnDel) btnDel.disabled = !on || ro;
        }

        function applyGmRestore() {
            if (opts.restoreScrollTop != null && wrap) {
                wrap.scrollTop = opts.restoreScrollTop;
            }
            var gmRows = tbody ? tbody.querySelectorAll("tr.gm-row") : [];
            if (opts.selectGroupId && tbody) {
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

        if (tbody) {
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
        }

        var bSel = document.getElementById("gmBtnSelect");
        if (bSel) {
            bSel.addEventListener("click", function () {
                var gid = selectedGroupId();
                if (!gid || !tbody) return;
                var tr = tbody.querySelector("tr.gm-row.selected");
                var idx = tr ? Array.prototype.indexOf.call(tbody.querySelectorAll("tr.gm-row"), tr) : -1;
                var st = wrap ? wrap.scrollTop : 0;
                openGroupSetup("2", gid, null, {
                    restoreSelectionIndex: idx,
                    restoreScrollTop: st
                });
            });
        }

        var bAdd = document.getElementById("gmBtnAdd");
        if (bAdd) {
            bAdd.addEventListener("click", function () {
                if (ro) return;
                var trSel = tbody ? tbody.querySelector("tr.gm-row.selected") : null;
                var idx = trSel
                    ? Array.prototype.indexOf.call(tbody.querySelectorAll("tr.gm-row"), trSel)
                    : -1;
                var st = wrap ? wrap.scrollTop : 0;
                openGroupSetup("1", null, null, {
                    restoreSelectionIndex: idx,
                    restoreScrollTop: st
                });
            });
        }

        var bDel = document.getElementById("gmBtnDelete");
        if (bDel) {
            bDel.addEventListener("click", function () {
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
                refreshGroupMaintenanceContent({ restoreScrollTop: st });
            });
        }

        applyGmRestore();
    }

    /**
     * frmGroupMaintenance — grid GroupID + Group Name; Select / Add New / Delete / Close (584×336).
     */
    function openGroupMaintenance(opts) {
        opts = opts || {};
        openAppModal(
            "Group Maintenance - Binventory Workstation",
            buildGroupMaintenanceBodyHtml(),
            "",
            "modal-group-maint modal-footer-hidden modal-win-toolwindow"
        );
        bindGroupMaintenanceEvents(opts);
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
        vesselGroupAssignCallback = null;
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

    function pad2Sched(n) {
        return n < 10 ? "0" + n : String(n);
    }

    function minutesFromHHMM(hhmm) {
        var p = String(hhmm || "0:0").split(":");
        var h = parseInt(p[0], 10);
        var m = parseInt(p[1], 10);
        if (isNaN(h)) h = 0;
        if (isNaN(m)) m = 0;
        return h * 60 + m;
    }

    function vesselIdsForSchedule(sch) {
        ensureScheduleFields(sch);
        var o = {};
        (sch.vesselIds || []).forEach(function (id) {
            o[id] = true;
        });
        (sch.groupIds || []).forEach(function (gid) {
            var g = findGroupById(gid);
            if (g && g.vesselIds) {
                g.vesselIds.forEach(function (id) {
                    o[id] = true;
                });
            }
        });
        return Object.keys(o);
    }

    function vesselAppliesToCurrentSite(v) {
        var sid = state.currentWorkstationSiteId;
        return v && (v.siteId == null || v.siteId === sid);
    }

    /**
     * True when this calendar minute should run (matches Measurement Schedule types 0–4 UI).
     * Run-once uses state._schedRunOnceFired so it fires a single time per schedule id.
     */
    function scheduleMatchesThisMinute(sch, now) {
        ensureScheduleFields(sch);
        var hm = pad2Sched(now.getHours()) + ":" + pad2Sched(now.getMinutes());
        var info = (sch.scheduleInfoTime || sch.time || "08:00").slice(0, 5);
        var st = sch.scheduleType != null ? sch.scheduleType : 1;

        if (st === 0) {
            var ymd =
                now.getFullYear() + "-" + pad2Sched(now.getMonth() + 1) + "-" + pad2Sched(now.getDate());
            var sd = String(sch.scheduleStartDate || "").slice(0, 10);
            if (ymd !== sd || hm !== info) return false;
            state._schedRunOnceFired = state._schedRunOnceFired || {};
            return !state._schedRunOnceFired[sch.id];
        }
        if (st === 1) {
            var dow = now.getDay();
            if (!sch.scheduleDays || !sch.scheduleDays[dow]) return false;
            var cur = now.getHours() * 60 + now.getMinutes();
            var start = minutesFromHHMM(sch.timeIntervalStart || "08:00");
            var end = minutesFromHHMM(sch.timeIntervalEnd || "17:00");
            if (cur < start || cur > end) return false;
            var iv = parseInt(sch.timeIntervalMinutes, 10) || 60;
            if (iv <= 0) return false;
            var rel = cur - start;
            return rel % iv === 0;
        }
        if (st === 2) {
            var dow2 = now.getDay();
            if (!sch.scheduleDays || !sch.scheduleDays[dow2]) return false;
            return hm === info;
        }
        if (st === 3) {
            if (hm !== info) return false;
            var sdStr = sch.scheduleStartDate || todayIsoDateLocal();
            var sd = new Date(sdStr + "T12:00:00");
            if (isNaN(sd.getTime())) return false;
            return now.getDay() === sd.getDay();
        }
        if (st === 4) {
            if (hm !== info) return false;
            var sdStr2 = sch.scheduleStartDate || todayIsoDateLocal();
            var parts = sdStr2.split("-");
            var dom = parseInt(parts[2], 10);
            if (isNaN(dom)) return false;
            return now.getDate() === dom;
        }
        return false;
    }

    function applyScheduledMeasurementForVessel(v) {
        var sid = parseInt(v.sensorTypeId, 10);
        if (sid === 14 || sid === 15) {
            applySimulatedMeasurementValues(v);
            v.status = defaultIdleStatusForSensorType(sid);
            v.lastMeasurement = new Date().toLocaleString();
            return;
        }
        measureVessel(v, { silent: true, scheduled: true });
    }

    /** Called when eBob Engine + Scheduler services are running (matches services.msc). */
    function processEmulatedScheduleTick(now) {
        if (!state.ebobServicesRunning || !state.ebobSchedulerRunning) return;
        var any = false;
        schedulesMeasurementList().forEach(function (sch) {
            ensureScheduleFields(sch);
            if (sch.scheduleActive === false || sch.draft) return;
            if (!scheduleMatchesThisMinute(sch, now)) return;
            var st = sch.scheduleType != null ? sch.scheduleType : 1;
            if (st === 0) {
                state._schedRunOnceFired = state._schedRunOnceFired || {};
                state._schedRunOnceFired[sch.id] = true;
            }
            vesselIdsForSchedule(sch).forEach(function (vid) {
                var v = findVessel(vid);
                if (!v || !vesselAppliesToCurrentSite(v)) return;
                applyScheduledMeasurementForVessel(v);
                any = true;
            });
        });
        if (any) {
            saveState();
            renderGrid();
        }
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
            '<input type="text" id="mssName" class="mss-input-name" maxlength="50" value="' +
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
            if (nm.length > 50) {
                toast("Schedule Name is required and must be 50 characters or less.");
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
            toast("Schedule saved.");
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
            toast("Temporary group applied to the vessel display.");
        });
    }

    function buildContactMaintenanceBodyHtml() {
        state.contacts.forEach(ensureContactRecord);
        var rowsHtml = state.contacts
            .map(function (c, idx) {
                var fn = String(c.firstName || "").trim();
                var ln = String(c.lastName || "").trim();
                var em = String(c.email || c.emailAddress || "").trim();
                return (
                    '<tr data-cm-idx="' +
                    idx +
                    '" tabindex="-1">' +
                    "<td>" +
                    escapeHtml(fn) +
                    "</td><td>" +
                    escapeHtml(ln) +
                    "</td><td>" +
                    escapeHtml(em) +
                    "</td></tr>"
                );
            })
            .join("");
        return (
            '<div class="cm-win">' +
            '<div class="cm-lbl-title">Contact Maintenance</div>' +
            '<div class="cm-main-row">' +
            '<div class="cm-dgv-wrap" role="presentation">' +
            '<table class="win-dgv cm-dgv" tabindex="0">' +
            "<thead><tr>" +
            "<th>First Name</th>" +
            "<th>Last Name</th>" +
            "<th>Email</th>" +
            "</tr></thead>" +
            "<tbody>" +
            rowsHtml +
            "</tbody></table></div>" +
            '<div class="cm-actions-col">' +
            '<button type="button" class="win-btn win-btn-default" id="cmSelect" disabled accesskey="s"><u>S</u>elect</button>' +
            '<button type="button" class="win-btn" id="cmAddNew" accesskey="a"><u>A</u>dd New</button>' +
            '<button type="button" class="win-btn" id="cmDelete" disabled accesskey="d"><u>D</u>elete</button>' +
            '<span class="cm-actions-spacer" aria-hidden="true"></span>' +
            '<button type="button" class="win-btn" data-close-app accesskey="c"><u>C</u>lose</button>' +
            "</div></div></div>"
        );
    }

    function refreshContactMaintenanceContent(opts) {
        opts = opts || {};
        if (!appModalShell.classList.contains("modal-contact-maint")) {
            openContactMaintenance(opts);
            return;
        }
        updateAppModalContent("Contact Maintenance - Binventory Workstation", buildContactMaintenanceBodyHtml(), "");
        bindContactMaintenanceEvents(opts);
    }

    function bindContactMaintenanceEvents(opts) {
        opts = opts || {};
        var root = appModalBody;
        if (!root) return;
        var tbody = root.querySelector(".cm-dgv tbody");
        var selBtn = root.querySelector("#cmSelect");
        var delBtn = root.querySelector("#cmDelete");
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
            if (!tbody || idx < 0 || idx >= state.contacts.length) {
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
            if (selectedIdx < 0 || selectedIdx >= state.contacts.length) return;
            openContactSetupDialog({ isNew: false, editIndex: selectedIdx });
        }

        if (tbody) {
            tbody.addEventListener("click", function (e) {
                var tr = e.target.closest("tr[data-cm-idx]");
                if (!tr) return;
                e.preventDefault();
                applySelection(parseInt(tr.getAttribute("data-cm-idx"), 10));
            });
            tbody.addEventListener("dblclick", function (e) {
                var tr = e.target.closest("tr[data-cm-idx]");
                if (!tr) return;
                e.preventDefault();
                applySelection(parseInt(tr.getAttribute("data-cm-idx"), 10));
                doOpenSelection();
            });
        }

        if (selBtn) selBtn.addEventListener("click", doOpenSelection);

        var addNew = root.querySelector("#cmAddNew");
        if (addNew) {
            addNew.addEventListener("click", function () {
                openContactSetupDialog({ isNew: true });
            });
        }

        if (delBtn) {
            delBtn.addEventListener("click", function () {
                if (selectedIdx < 0 || selectedIdx >= state.contacts.length) return;
                var c = state.contacts[selectedIdx];
                var nm = String(c.name || (String(c.firstName || "") + " " + String(c.lastName || "")).trim() || c.id);
                showBinventoryMessageBox({
                    icon: "warn",
                    message: "Are you sure you want to delete '" + nm + "'?",
                    buttons: "okcancel",
                    onOk: function () {
                        var cid = c.id;
                        removeContactIdFromAllVessels(cid);
                        state.contacts.splice(selectedIdx, 1);
                        saveState();
                        closeAppModal();
                        openContactMaintenance();
                    }
                });
            });
        }

        clearSelection();
        if (opts.selectIndex != null && opts.selectIndex >= 0 && opts.selectIndex < state.contacts.length) {
            applySelection(opts.selectIndex);
            var tr = tbody && tbody.querySelector('tr[data-cm-idx="' + opts.selectIndex + '"]');
            if (tr) tr.scrollIntoView({ block: "nearest" });
        }
    }

    /**
     * frmContactSetup — ClientSize 485×173. Stacks on #backdropAppStack over Contact Maintenance (ShowDialog).
     */
    function openContactSetupDialog(opts) {
        opts = opts || {};
        var isNew = opts.isNew === true;
        var editIndex = opts.editIndex;
        if (!isNew && (editIndex == null || editIndex < 0 || editIndex >= state.contacts.length)) return;
        var c = isNew ? {} : state.contacts[editIndex];
        if (!isNew) ensureContactRecord(c);

        var fn = isNew ? "" : String(c.firstName || "");
        var ln = isNew ? "" : String(c.lastName || "");
        var em = isNew ? "" : String(c.email || c.emailAddress || "");
        var job = isNew ? "" : String(c.jobTitle || "");

        var html =
            '<div class="cs-binv">' +
            '<div class="cs-binv-title">Contact Setup</div>' +
            '<div class="cs-binv-rows">' +
            "<label class=\"cs-row\"><span>First Name:</span><input type=\"text\" id=\"csFirstName\" class=\"cs-inp\" maxlength=\"80\" value=\"" +
            escapeHtml(fn) +
            '"></label>' +
            "<label class=\"cs-row\"><span>Last Name:</span><input type=\"text\" id=\"csLastName\" class=\"cs-inp\" maxlength=\"80\" value=\"" +
            escapeHtml(ln) +
            '"></label>' +
            "<label class=\"cs-row\"><span>Email Address:</span><input type=\"text\" id=\"csEmail\" class=\"cs-inp\" maxlength=\"120\" value=\"" +
            escapeHtml(em) +
            '"></label>' +
            "<label class=\"cs-row\"><span>Job Title:</span><input type=\"text\" id=\"csJobTitle\" class=\"cs-inp\" maxlength=\"120\" value=\"" +
            escapeHtml(job) +
            '"></label>' +
            "</div>" +
            '<div class="cs-binv-actions">' +
            '<button type="button" class="win-btn win-btn-default" id="csSave"><u>S</u>ave</button>' +
            '<button type="button" class="win-btn" id="csCancel"><u>C</u>ancel</button>' +
            "</div></div>";

        openStackedAppModal("Contact Setup - Binventory Workstation", html, "", "modal-contact-setup modal-footer-hidden modal-win-toolwindow");

        document.getElementById("csCancel").addEventListener("click", function () {
            closeStackedAppModal();
        });

        document.getElementById("csSave").addEventListener("click", function () {
            var f = (document.getElementById("csFirstName") && document.getElementById("csFirstName").value.trim()) || "";
            var l = (document.getElementById("csLastName") && document.getElementById("csLastName").value.trim()) || "";
            var mail = (document.getElementById("csEmail") && document.getElementById("csEmail").value.trim()) || "";
            var jt = (document.getElementById("csJobTitle") && document.getElementById("csJobTitle").value.trim()) || "";
            var disp = (f + " " + l).trim() || (mail || "Contact");

            if (isNew) {
                state.contacts.push({
                    id: uid("c"),
                    firstName: f,
                    lastName: l,
                    name: disp,
                    email: mail,
                    emailAddress: mail,
                    phone: "",
                    jobTitle: jt
                });
                ensureContactRecord(state.contacts[state.contacts.length - 1]);
            } else {
                var tgt = state.contacts[editIndex];
                tgt.firstName = f;
                tgt.lastName = l;
                tgt.name = disp;
                tgt.email = mail;
                tgt.emailAddress = mail;
                tgt.jobTitle = jt;
                ensureContactRecord(tgt);
            }
            saveState();
            var selIdx = isNew ? state.contacts.length - 1 : editIndex;
            closeStackedAppModal();
            refreshContactMaintenanceContent({ selectIndex: selIdx });
            toast("Contact saved.");
        });

        setTimeout(function () {
            var el = document.getElementById("csFirstName");
            if (el) el.focus();
        }, 0);
    }

    /**
     * frmContactMaintenance — ClientSize 584×336 (dgv 451×288 at 12,38; Select/Add/Delete/Close on right).
     */
    function openContactMaintenance(opts) {
        opts = opts || {};
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to open Contact Maintenance.");
            return;
        }
        openAppModal(
            "Contact Maintenance - Binventory Workstation",
            buildContactMaintenanceBodyHtml(),
            "",
            "modal-contact-maint modal-footer-hidden modal-win-toolwindow"
        );
        bindContactMaintenanceEvents(opts);
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

    function buildSiteMaintenanceBodyHtml() {
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

        return (
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
            "</div>"
        );
    }

    function refreshSiteMaintenanceContent(opts) {
        opts = opts || {};
        if (!appModalShell.classList.contains("modal-site-maintenance")) {
            openSiteMaintenance(opts);
            return;
        }
        if (opts.selectSiteId != null) smSelectedSiteId = opts.selectSiteId;
        updateAppModalContent(
            "Site Maintenance - Binventory Workstation",
            buildSiteMaintenanceBodyHtml(),
            ""
        );
        bindSiteMaintenanceEvents();
    }

    function bindSiteMaintenanceEvents() {
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
                refreshSiteMaintenanceContent({});
                toast("Site deleted.");
            });
        }

        document.getElementById("smAddNew").addEventListener("click", function () {
            openSiteSetupModal(true);
        });
    }

    function openSiteSetupModal(isNew, siteId) {
        var site = isNew ? null : findSite(siteId);
        if (!isNew && !site) {
            closeStackedAppModal();
            refreshSiteMaintenanceContent({});
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

        openStackedAppModal("Site Setup - Binventory Workstation", html, footer, "modal-site-setup");

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
            closeStackedAppModal();
            refreshSiteMaintenanceContent({ selectSiteId: isNew ? draftId : siteId });
        });
        document.getElementById("ssuCancel").addEventListener("click", function () {
            closeStackedAppModal();
            refreshSiteMaintenanceContent({});
        });
    }

    function openSiteMaintenance(opts) {
        opts = opts || {};
        state.sites.forEach(ensureSiteFields);
        saveState();

        if (opts.selectSiteId != null) smSelectedSiteId = opts.selectSiteId;
        else smSelectedSiteId = null;

        openAppModal(
            "Site Maintenance - Binventory Workstation",
            buildSiteMaintenanceBodyHtml(),
            "",
            "modal-site-maintenance"
        );
        bindSiteMaintenanceEvents();
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

    /** Seconds — default Status Request Delay for Protocol A (frmSensorNetworkSetup / BLL). Not a display hack; the stored value is numeric. */
    var EBOB_STATUS_REQUEST_DELAY_DEFAULT = 1;

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
        if (!n.interface) n.interface = "COM2";
        if (!n.commParams) n.commParams = "2400,8,N,1";
        if (!n.name) n.name = "COM2 Sensor Network";
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
        if (n.statusRequestDelay == null) n.statusRequestDelay = EBOB_STATUS_REQUEST_DELAY_DEFAULT;
        parseCommParamsToFields(n);
        if (n.siteId == null && state.sites && state.sites.length) {
            n.siteId = state.sites[0].id;
        }
    }

    /**
     * WinForms-style NumericUpDown for Sensor Network Advanced (not type=number).
     * Keys = input id; up/down = button ids.
     */
    var SNS_NUD_FIELDS = {
        snsAtt: {
            up: "snsAttUp",
            down: "snsAttDown",
            step: 1,
            min: 1,
            max: 10,
            decimals: 0,
            defaultVal: 3
        },
        snsStatusDelay: {
            up: "snsDelayUp",
            down: "snsDelayDown",
            step: 0.1,
            min: 0,
            max: 5,
            decimals: 1,
            defaultVal: EBOB_STATUS_REQUEST_DELAY_DEFAULT
        },
        /* Drop / Transmit: frmSensorNetworkSetup.Designer.vb Increment = 5 */
        snsDropTimeout: {
            up: "snsDropUp",
            down: "snsDropDown",
            step: 5,
            min: 20,
            max: 200,
            decimals: 0,
            defaultVal: 120
        },
        snsTransmitDelay: {
            up: "snsTransmitUp",
            down: "snsTransmitDown",
            step: 5,
            min: 5,
            max: 500,
            decimals: 0,
            defaultVal: 20
        },
        /* numResponseTimeout: Increment 0.1, Minimum 0.5, Maximum 20 */
        snsRto: {
            up: "snsRtoUp",
            down: "snsRtoDown",
            step: 0.1,
            min: 0.5,
            max: 20,
            decimals: 1,
            defaultVal: 1.5
        }
    };

    var SNS_NUD_SVG_UP =
        '<svg width="9" height="6" viewBox="0 0 9 6" aria-hidden="true" focusable="false"><path fill="currentColor" d="M4.5 0L9 6H0z"/></svg>';
    var SNS_NUD_SVG_DOWN =
        '<svg width="9" height="6" viewBox="0 0 9 6" aria-hidden="true" focusable="false"><path fill="currentColor" d="M4.5 6L9 0H0z"/></svg>';

    function snsNudClampVal(cfg, x) {
        if (!isFinite(x)) {
            return cfg.defaultVal != null ? cfg.defaultVal : cfg.min;
        }
        return Math.min(cfg.max, Math.max(cfg.min, x));
    }

    /** Snap to min + n*step (frmSensorNetworkSetup NumericUpDown Increment). */
    function snsNudSnapIntStep(cfg, v) {
        v = snsNudClampVal(cfg, v);
        if (cfg.decimals !== 0) return v;
        if (!cfg.step || cfg.step <= 1) {
            return Math.round(v);
        }
        var q = Math.round((v - cfg.min) / cfg.step);
        var snapped = cfg.min + q * cfg.step;
        return snsNudClampVal(cfg, snapped);
    }

    function snsNudFormat(cfg, x) {
        x = snsNudClampVal(cfg, x);
        if (cfg.decimals === 0) {
            x = snsNudSnapIntStep(cfg, x);
            return String(Math.round(x));
        }
        return x.toFixed(1);
    }

    function snsNudParseRaw(cfg, str) {
        var raw = parseFloat(String(str).trim().replace(/,/g, "."));
        if (!isFinite(raw)) {
            return cfg.defaultVal != null ? cfg.defaultVal : cfg.min;
        }
        return snsNudClampVal(cfg, raw);
    }

    function snsNudParse(cfg, str) {
        var raw = parseFloat(String(str).trim().replace(/,/g, "."));
        if (!isFinite(raw)) {
            return cfg.defaultVal != null ? cfg.defaultVal : cfg.min;
        }
        var v = snsNudClampVal(cfg, raw);
        if (cfg.decimals === 0) {
            return snsNudSnapIntStep(cfg, v);
        }
        return v;
    }

    function snsNudRefreshButtons(inputId) {
        var cfg = SNS_NUD_FIELDS[inputId];
        if (!cfg) return;
        var inp = document.getElementById(inputId);
        var up = document.getElementById(cfg.up);
        var down = document.getElementById(cfg.down);
        if (!inp || !up || !down) return;
        var v = snsNudParse(cfg, inp.value);
        if (cfg.decimals === 0) {
            up.disabled = v >= cfg.max;
            down.disabled = v <= cfg.min;
        } else {
            up.disabled = v >= cfg.max - 1e-9;
            down.disabled = v <= cfg.min + 1e-9;
        }
    }

    function snsNudApplyStep(inputId, sign) {
        var cfg = SNS_NUD_FIELDS[inputId];
        if (!cfg) return;
        var el = document.getElementById(inputId);
        if (!el) return;
        var cur = snsNudParseRaw(cfg, el.value);
        var v = cur + sign * cfg.step;
        if (cfg.decimals === 0) {
            v = snsNudSnapIntStep(cfg, v);
        } else {
            v = Math.round(v * 10) / 10;
            v = snsNudClampVal(cfg, v);
        }
        el.value = snsNudFormat(cfg, v);
        snsNudRefreshButtons(inputId);
    }

    function snsNudCommitBlur(inputId) {
        var cfg = SNS_NUD_FIELDS[inputId];
        var el = document.getElementById(inputId);
        if (!cfg || !el) return;
        el.value = snsNudFormat(cfg, snsNudParse(cfg, el.value));
        snsNudRefreshButtons(inputId);
    }

    function refreshAllSensorNetworkNudButtons() {
        Object.keys(SNS_NUD_FIELDS).forEach(function (id) {
            snsNudRefreshButtons(id);
        });
    }

    function snsNumericUpDownControlHtml(inputId, valueStr, cfg, titles) {
        var t = titles || {};
        var tu = t.up || "Increase";
        var td = t.down || "Decrease";
        return (
            '<span class="sns-numeric-updown" role="group">' +
            '<input type="text" inputmode="decimal" id="' +
            inputId +
            '" class="sns-nud-input" autocomplete="off" value="' +
            escapeHtml(valueStr) +
            '">' +
            '<span class="sns-nud-spin">' +
            '<button type="button" class="sns-nud-btn sns-nud-btn-up" id="' +
            cfg.up +
            '" tabindex="-1" title="' +
            escapeHtml(tu) +
            '" aria-label="' +
            escapeHtml(tu) +
            '">' +
            SNS_NUD_SVG_UP +
            "</button>" +
            '<button type="button" class="sns-nud-btn sns-nud-btn-down" id="' +
            cfg.down +
            '" tabindex="-1" title="' +
            escapeHtml(td) +
            '" aria-label="' +
            escapeHtml(td) +
            '">' +
            SNS_NUD_SVG_DOWN +
            "</button>" +
            "</span>" +
            "</span>"
        );
    }

    function wireSensorNetworkAdvancedNumericUpDowns() {
        Object.keys(SNS_NUD_FIELDS).forEach(function (inputId) {
            var cfg = SNS_NUD_FIELDS[inputId];
            var inp = document.getElementById(inputId);
            var up = document.getElementById(cfg.up);
            var down = document.getElementById(cfg.down);
            if (!inp || !up || !down) return;
            up.addEventListener("click", function (e) {
                e.preventDefault();
                snsNudApplyStep(inputId, 1);
            });
            down.addEventListener("click", function (e) {
                e.preventDefault();
                snsNudApplyStep(inputId, -1);
            });
            inp.addEventListener("blur", function () {
                snsNudCommitBlur(inputId);
            });
            snsNudRefreshButtons(inputId);
        });
    }

    function formatStatusRequestDelayForDisplay(seconds) {
        return snsNudFormat(SNS_NUD_FIELDS.snsStatusDelay, seconds);
    }

    function parseStatusRequestDelayField(str) {
        return snsNudParse(SNS_NUD_FIELDS.snsStatusDelay, str);
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
        "Sugar B 1",
        "Sugar B 2",
        "Grain A",
        "Grain B",
        "Pellets",
        "Bulk Store"
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
                name: "COM2 Sensor Network",
                protocol: "Protocol A",
                interface: "COM2",
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
                remotePort: "50000",
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
                interface: "COM2",
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
        resolveDuplicateSensorAddressesOnAllVessels();
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
                if (att) att.value = snsNudFormat(SNS_NUD_FIELDS.snsAtt, 3);
                if (rto) rto.value = snsNudFormat(SNS_NUD_FIELDS.snsRto, 2);
                if (drop) drop.value = snsNudFormat(SNS_NUD_FIELDS.snsDropTimeout, 120);
                if (stk) stk.checked = false;
                break;
            case "Protocol A":
                setSel(baud, 2400);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = snsNudFormat(SNS_NUD_FIELDS.snsAtt, 3);
                if (rto) rto.value = snsNudFormat(SNS_NUD_FIELDS.snsRto, 1.5);
                if (std) std.value = snsNudFormat(SNS_NUD_FIELDS.snsStatusDelay, EBOB_STATUS_REQUEST_DELAY_DEFAULT);
                break;
            case "Modbus/RTU":
                setSel(baud, 9600);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = snsNudFormat(SNS_NUD_FIELDS.snsAtt, 3);
                if (rto) rto.value = snsNudFormat(SNS_NUD_FIELDS.snsRto, 3.5);
                if (txd) txd.value = snsNudFormat(SNS_NUD_FIELDS.snsTransmitDelay, 20);
                break;
            case "SPL-100 Push":
            case "SPL-200 Push":
                setSel(baud, 9600);
                setSel(db, 8);
                setSel(par, "None");
                setSel(stb, 1);
                if (att) att.value = snsNudFormat(SNS_NUD_FIELDS.snsAtt, 3);
                if (rto) rto.value = snsNudFormat(SNS_NUD_FIELDS.snsRto, 1);
                break;
            case "HART Protocol":
                setSel(baud, 1200);
                setSel(db, 8);
                setSel(par, "Odd");
                setSel(stb, 1);
                if (att) att.value = snsNudFormat(SNS_NUD_FIELDS.snsAtt, 3);
                if (rto) rto.value = snsNudFormat(SNS_NUD_FIELDS.snsRto, 2);
                if (txd) txd.value = snsNudFormat(SNS_NUD_FIELDS.snsTransmitDelay, 60);
                break;
            default:
                break;
        }
        refreshAllSensorNetworkNudButtons();
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
        var cellStuck = document.getElementById("snsAdvCellStuck");
        var advWarn = document.getElementById("snsAdvWarn");

        var spl = p === "SPL-100 Push" || p === "SPL-200 Push";
        if (advGrp) advGrp.style.display = spl ? "none" : "";
        if (advWarn) advWarn.style.display = spl ? "none" : "";
        var snShell = document.getElementById("snSetupShell");
        if (snShell) snShell.classList.toggle("sns-compact", spl);

        /* Advanced extras per frmSensorNetworkSetup.vb cboProtocol_SelectedIndexChanged */
        if (rowStd) rowStd.style.display = p === "Protocol A" ? "" : "none";
        if (rowDrop) rowDrop.style.display = p === "Protocol B" ? "" : "none";
        if (rowTx) rowTx.style.display = p === "Modbus/RTU" || p === "HART Protocol" ? "" : "none";
        if (cellStuck) cellStuck.style.display = p === "Protocol B" ? "" : "none";

        refreshAllSensorNetworkNudButtons();

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
            '<fieldset class="sns-group sns-group-protocol" id="snsGrpProto"><legend>Protocol</legend>' +
            '<div class="sns-proto-row">' +
            '<label class="sns-proto-lbl" for="snsProtocol">Protocol used by all sensors on this network:</label>' +
            '<select id="snsProtocol" class="sns-input-protocol">' +
            buildProtocolOptionsHtml(n.protocol) +
            "</select>" +
            "</div>" +
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
            '<div class="sns-adv-grid">' +
            '<div class="sns-adv-cell">' +
            '<label class="sns-nud-lbl"><span>Request Attempts:</span>' +
            snsNumericUpDownControlHtml(
                "snsAtt",
                snsNudFormat(SNS_NUD_FIELDS.snsAtt, n.requestAttempts),
                SNS_NUD_FIELDS.snsAtt,
                { up: "Increase by 1", down: "Decrease by 1" }
            ) +
            "</label>" +
            "</div>" +
            '<div class="sns-adv-cell sns-adv-cell-top-right" id="snsAdvTopRight">' +
            '<div class="sns-adv-extra" id="snsRowStatusDelay">' +
            '<label class="sns-nud-lbl"><span>Status Request Delay:</span>' +
            snsNumericUpDownControlHtml(
                "snsStatusDelay",
                formatStatusRequestDelayForDisplay(n.statusRequestDelay),
                SNS_NUD_FIELDS.snsStatusDelay,
                { up: "Increase by 0.1 s", down: "Decrease by 0.1 s" }
            ) +
            "</label>" +
            "</div>" +
            '<div class="sns-adv-extra" id="snsRowDrop">' +
            '<label class="sns-nud-lbl"><span>Drop Timeout:</span>' +
            snsNumericUpDownControlHtml(
                "snsDropTimeout",
                snsNudFormat(SNS_NUD_FIELDS.snsDropTimeout, n.dropTimeout),
                SNS_NUD_FIELDS.snsDropTimeout,
                { up: "Increase by 5", down: "Decrease by 5" }
            ) +
            "</label>" +
            "</div>" +
            '<div class="sns-adv-extra" id="snsRowTransmit">' +
            '<label class="sns-nud-lbl"><span>Transmit Delay:</span>' +
            snsNumericUpDownControlHtml(
                "snsTransmitDelay",
                snsNudFormat(SNS_NUD_FIELDS.snsTransmitDelay, n.transmitDelay),
                SNS_NUD_FIELDS.snsTransmitDelay,
                { up: "Increase by 5", down: "Decrease by 5" }
            ) +
            "</label>" +
            "</div>" +
            "</div>" +
            '<div class="sns-adv-cell">' +
            '<label class="sns-nud-lbl"><span>Response Timeout:</span>' +
            snsNumericUpDownControlHtml(
                "snsRto",
                snsNudFormat(SNS_NUD_FIELDS.snsRto, n.responseTimeout),
                SNS_NUD_FIELDS.snsRto,
                { up: "Increase by 0.1 s", down: "Decrease by 0.1 s" }
            ) +
            "</label>" +
            "</div>" +
            '<div class="sns-adv-cell sns-adv-cell-stuck" id="snsAdvCellStuck">' +
            '<div class="sns-adv-extra sns-adv-stuck">' +
            "<label class=\"sns-stuck-lbl\"><span>Disable StuckTop:</span>" +
            '<input type="checkbox" id="snsDisableStuckTop"' +
            (n.disableStuckTop ? " checked" : "") +
            "></label>" +
            "</div>" +
            "</div>" +
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
        wireSensorNetworkAdvancedNumericUpDowns();
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
        var attEl = document.getElementById("snsAtt");
        var rtoEl = document.getElementById("snsRto");
        var stdEl = document.getElementById("snsStatusDelay");
        var dropEl = document.getElementById("snsDropTimeout");
        var txEl = document.getElementById("snsTransmitDelay");
        var stkEl = document.getElementById("snsDisableStuckTop");
        if (attEl) n.requestAttempts = Math.round(snsNudParse(SNS_NUD_FIELDS.snsAtt, attEl.value));
        if (rtoEl) n.responseTimeout = snsNudParse(SNS_NUD_FIELDS.snsRto, rtoEl.value);
        if (n.protocol === "Protocol A" && stdEl) {
            n.statusRequestDelay = snsNudParse(SNS_NUD_FIELDS.snsStatusDelay, stdEl.value);
        }
        if (dropEl) n.dropTimeout = Math.round(snsNudParse(SNS_NUD_FIELDS.snsDropTimeout, dropEl.value));
        if (txEl) n.transmitDelay = snsNudParse(SNS_NUD_FIELDS.snsTransmitDelay, txEl.value);
        if (isNaN(n.transmitDelay)) n.transmitDelay = 20;
        if (stkEl) n.disableStuckTop = stkEl.checked;
        syncCommParamsString(n);
        saveState();
        closeSensorNetworkSetup();
        toast("Sensor network saved.");
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
                toast("Network deleted.");
            });
        }
        var btnDev = document.getElementById("snBtnDevice");
        if (btnDev) {
            btnDev.addEventListener("click", function () {
                if (snSelectedNetworkId) openSensorNetworkSetup(snSelectedNetworkId);
            });
        }

        document.getElementById("snAddNew").addEventListener("click", function () {
            /* Default serial port is COM2; additional networks use COM3, COM4, … */
            var comNum = Math.min(6, state.sensorNetworks.length + 2);
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
            toast("Network added.");
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
                '<button type="button" class="win-btn win-btn-default" data-close-app accesskey="c">Cancel</button>',
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
                toast("Email settings saved.");
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

    /**
     * When "Do NOT require users to log in" is on (autoLogin), session matches eBob after LoadApplication
     * (no frmLogin) — gsLoggedInUserID effective as admin until Logoff.
     */
    function applyAutoLoginSessionIfEnabled() {
        ensureSystemSettingsDefaults();
        if (state.systemSettings.autoLogin !== true) {
            return;
        }
        var adm = state.users.filter(function (u) {
            return String(u.userId).toLowerCase() === "admin";
        })[0];
        if (adm) {
            state.currentUser = adm.name;
        } else {
            state.currentUser = null;
        }
        syncMenuStripForSession();
        updateTitleBar();
        updateMeasureMenuDisabled();
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
            if (s.autoLogin) {
                applyAutoLoginSessionIfEnabled();
            }
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
            toast("Password updated.");
        });
    }

    /**
     * frmUserMaintenance grid User Type — same strings as frmUserSetup cboSecurityAccess (BLL access levels).
     */
    function userTypeLabelFromAccessLevel(level) {
        var n = parseInt(level, 10);
        if (n === 1) return "Administrator User";
        if (n === 3) return "System User (Read-Only)";
        return "System User";
    }

    function accessLevelFromSecurityComboLabel(label) {
        if (label === "Administrator User") return 1;
        if (label === "System User") return 2;
        return 3;
    }

    /**
     * frmUserSetup / SecurityRecord — accessLevel, authenticationMethod, password, jobTitle.
     */
    function ensureUserRecord(u) {
        if (!u || typeof u !== "object") return u;
        if (u.accessLevel == null) {
            if (u.role === "Administrator" || u.userType === "Administrator" || u.userType === "Administrator User") {
                u.accessLevel = 1;
            } else if (u.userType === "System User (Read-Only)" || u.role === "Read Only") {
                u.accessLevel = 3;
            } else {
                u.accessLevel = 2;
            }
        }
        if (u.authenticationMethod == null) u.authenticationMethod = 0;
        if (u.jobTitle == null) u.jobTitle = "";
        if (u.password == null) u.password = "";
        u.userType = userTypeLabelFromAccessLevel(u.accessLevel);
        return u;
    }

    /**
     * frmUserMaintenance.vb — DataGridView columns: User ID, User Type, FirstName, Middle Init, Last Name, Last Logon;
     * buttons top→bottom: Select, Add New, Delete; Close at bottom (Designer ClientSize 859×336).
     */
    function ensureUserMaintenanceUser(u) {
        if (!u || typeof u !== "object") return u;
        if (u.userId == null || String(u.userId).trim() === "") u.userId = u.id || u.name || "User";
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
        ensureUserRecord(u);
        return u;
    }

    function loggedInUserIdString() {
        var rec = state.users.filter(function (x) {
            return x.name === state.currentUser || String(x.userId) === state.currentUser;
        })[0];
        if (rec && rec.userId != null) return String(rec.userId);
        return state.currentUser ? String(state.currentUser) : "";
    }

    /** Body HTML for frmUserMaintenance (primary #appModalShell only). */
    function buildUserMaintenanceBodyHtml() {
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
        return (
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
            "</div></div></div>"
        );
    }

    /**
     * Rebuild User Maintenance under #appModalShell while frmUserSetup (#backdropAppStack) is closed.
     */
    function refreshUserMaintenanceContent(opts) {
        opts = opts || {};
        if (!appModalShell.classList.contains("modal-user-maintenance")) {
            openUserMaintenance(opts);
            return;
        }
        updateAppModalContent("User Maintenance - Binventory Workstation", buildUserMaintenanceBodyHtml(), "");
        bindUserMaintenanceEvents(opts);
    }

    function bindUserMaintenanceEvents(opts) {
        opts = opts || {};
        var root = appModalBody;
        if (!root) return;
        var tbody = root.querySelector(".um-dgv tbody");
        var selBtn = root.querySelector("#umSelect");
        var delBtn = root.querySelector("#umDelete");
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
            openUserSetupDialog({ isNew: false, editIndex: selectedIdx });
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

        if (selBtn) selBtn.addEventListener("click", doOpenSelection);

        var addNew = root.querySelector("#umAddNew");
        if (addNew) {
            addNew.addEventListener("click", function () {
                openUserSetupDialog({ isNew: true });
            });
        }

        if (delBtn) {
            delBtn.addEventListener("click", function () {
                if (selectedIdx < 0 || selectedIdx >= state.users.length) return;
                var u = state.users[selectedIdx];
                var uidStr = String(u.userId);
                if (uidStr === loggedInUserIdString()) {
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
                    message: "Are you sure you want to delete the user '" + uidStr + "'?",
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
        if (opts.selectIndex != null && opts.selectIndex >= 0 && opts.selectIndex < state.users.length) {
            applySelection(opts.selectIndex);
            var tr = tbody && tbody.querySelector('tr[data-um-idx="' + opts.selectIndex + '"]');
            if (tr) tr.scrollIntoView({ block: "nearest" });
        }
    }

    /**
     * frmUserSetup — Add New (msAction=1) or edit (msAction=2). Stacks on #backdropAppStack over User Maintenance (ShowDialog).
     */
    function openUserSetupDialog(opts) {
        opts = opts || {};
        var isNew = opts.isNew === true;
        var editIndex = opts.editIndex;
        var u = isNew ? {} : state.users[editIndex];
        if (!isNew && (editIndex == null || editIndex < 0 || editIndex >= state.users.length)) return;
        if (!isNew) {
            u = state.users[editIndex];
            ensureUserMaintenanceUser(u);
        }

        var origUserId = isNew ? "" : String(u.userId);
        var secLabels = ["System User (Read-Only)", "System User", "Administrator User"];
        var secSel = userTypeLabelFromAccessLevel(isNew ? 3 : u.accessLevel);
        if (secLabels.indexOf(secSel) < 0) secSel = "System User";

        var authIdx = isNew ? 0 : u.authenticationMethod || 0;
        var pwdVal = isNew ? "" : String(u.password != null ? u.password : "");
        /* frmUserSetup: txtVerifyPassword is only filled when editing a *different* user
         * (gsLoggedInUserID <> msUserID). Editing your own profile leaves verify blank. */
        var verifyVal = "";
        if (!isNew) {
            var loggedId = loggedInUserIdString();
            var editingDifferentUser =
                !loggedId ||
                String(loggedId).toLowerCase() !== String(origUserId).toLowerCase();
            verifyVal = editingDifferentUser ? pwdVal : "";
        }
        var isAdminLock = String(u.userId || "").toLowerCase() === "admin";

        var secOpts = secLabels
            .map(function (lbl) {
                return (
                    "<option value=\"" +
                    escapeHtml(lbl) +
                    "\"" +
                    (lbl === secSel ? " selected" : "") +
                    ">" +
                    escapeHtml(lbl) +
                    "</option>"
                );
            })
            .join("");

        var html =
            '<div class="us-binv">' +
            '<div class="us-binv-title">User Setup</div>' +
            '<fieldset class="us-binv-group us-binv-gb-user"><legend>User Information</legend>' +
            '<div class="us-binv-grid us-binv-grid-user">' +
            "<label><span>First Name:</span><input type=\"text\" id=\"usFirstName\" class=\"us-binv-inp\" maxlength=\"80\" value=\"" +
            escapeHtml(isNew ? "" : u.firstName) +
            '"></label>' +
            "<label><span>Middle Init.</span><input type=\"text\" id=\"usMiddleInit\" class=\"us-binv-inp us-binv-mi\" maxlength=\"1\" value=\"" +
            escapeHtml(isNew ? "" : u.middleInit) +
            '"></label>' +
            "<label><span>Last Name:</span><input type=\"text\" id=\"usLastName\" class=\"us-binv-inp\" maxlength=\"80\" value=\"" +
            escapeHtml(isNew ? "" : u.lastName) +
            '"></label>' +
            "<label><span>User ID:</span><input type=\"text\" id=\"usUserId\" class=\"us-binv-inp" +
            (isAdminLock ? " us-binv-inp-locked" : "") +
            "\" maxlength=\"64\" value=\"" +
            escapeHtml(isNew ? "" : u.userId) +
            '"' +
            (isAdminLock ? " disabled" : "") +
            "></label>" +
            "<label><span>Job Title:</span><input type=\"text\" id=\"usJobTitle\" class=\"us-binv-inp\" maxlength=\"120\" value=\"" +
            escapeHtml(isNew ? "" : u.jobTitle || "") +
            '"></label>' +
            "</div></fieldset>" +
            '<fieldset class="us-binv-group us-binv-gb-auth"><legend>User Authentication</legend>' +
            '<div class="us-binv-grid us-binv-grid-auth">' +
            "<label><span>Security Access:</span><select id=\"usSecurity\" class=\"us-binv-sel\"" +
            (isAdminLock ? " disabled" : "") +
            ">" +
            secOpts +
            "</select></label>" +
            "<label><span>Authentication Method:</span><select id=\"usAuthMethod\" class=\"us-binv-sel\">" +
            '<option value="0"' +
            (authIdx === 0 ? " selected" : "") +
            ">Use Password Authentication</option>" +
            '<option value="1"' +
            (authIdx === 1 ? " selected" : "") +
            ">Use LDAP Authentication</option>" +
            "</select></label>" +
            "<label id=\"usLblPw\"><span>Password:</span><input type=\"password\" id=\"usPassword\" class=\"us-binv-inp\" autocomplete=\"new-password\" value=\"" +
            escapeHtml(pwdVal) +
            '"></label>' +
            "<label id=\"usLblVpw\"><span>Verify Password:</span><input type=\"password\" id=\"usVerifyPassword\" class=\"us-binv-inp\" autocomplete=\"new-password\" value=\"" +
            escapeHtml(verifyVal) +
            '"></label>' +
            "</div></fieldset>" +
            '<div class="us-binv-actions">' +
            '<button type="button" class="win-btn win-btn-default" id="usSave"><u>S</u>ave</button>' +
            '<button type="button" class="win-btn" id="usCancel"><u>C</u>ancel</button>' +
            "</div></div>";

        openStackedAppModal("User Setup - Binventory Workstation", html, "", "modal-user-setup modal-footer-hidden modal-win-toolwindow");

        function syncUsAuthPwVisibility() {
            var m = document.getElementById("usAuthMethod");
            var ldap = m && m.value === "1";
            ["usLblPw", "usLblVpw", "usPassword", "usVerifyPassword"].forEach(function (id) {
                var el = document.getElementById(id);
                if (!el) return;
                if (id === "usLblPw" || id === "usLblVpw") {
                    el.classList.toggle("us-binv-disabled", ldap);
                } else {
                    el.disabled = ldap;
                }
            });
        }

        syncUsAuthPwVisibility();
        var authEl = document.getElementById("usAuthMethod");
        if (authEl) authEl.addEventListener("change", syncUsAuthPwVisibility);

        document.getElementById("usCancel").addEventListener("click", function () {
            closeStackedAppModal();
        });

        document.getElementById("usSave").addEventListener("click", function () {
            var fn = (document.getElementById("usFirstName") && document.getElementById("usFirstName").value.trim()) || "";
            var mi = (document.getElementById("usMiddleInit") && document.getElementById("usMiddleInit").value.trim()) || "";
            var ln = (document.getElementById("usLastName") && document.getElementById("usLastName").value.trim()) || "";
            var uidEl = document.getElementById("usUserId");
            var userIdStr = uidEl && !uidEl.disabled ? uidEl.value.trim() : origUserId;
            var job = (document.getElementById("usJobTitle") && document.getElementById("usJobTitle").value.trim()) || "";
            var secEl = document.getElementById("usSecurity");
            var secLabel = secEl && !secEl.disabled ? secEl.value : userTypeLabelFromAccessLevel(u.accessLevel);
            var authM = parseInt(document.getElementById("usAuthMethod").value, 10) || 0;
            var pw = document.getElementById("usPassword") ? document.getElementById("usPassword").value : "";
            var vpw = document.getElementById("usVerifyPassword") ? document.getElementById("usVerifyPassword").value : "";

            if (!userIdStr) {
                showBinventoryMessageBox({ icon: "warn", message: "User ID is required.", buttons: "ok" });
                return;
            }
            var dup = state.users.some(function (x, i) {
                if (isNew) return String(x.userId).toLowerCase() === userIdStr.toLowerCase();
                return i !== editIndex && String(x.userId).toLowerCase() === userIdStr.toLowerCase();
            });
            if (dup) {
                showBinventoryMessageBox({ icon: "warn", message: "That User ID is already in use.", buttons: "ok" });
                return;
            }
            if (authM === 0) {
                if (pw !== vpw) {
                    showBinventoryMessageBox({ icon: "warn", message: "Password and Verify Password do not match.", buttons: "ok" });
                    return;
                }
                if (isNew && pw === "") {
                    showBinventoryMessageBox({ icon: "warn", message: "Password is required when using password authentication.", buttons: "ok" });
                    return;
                }
            }

            var al = accessLevelFromSecurityComboLabel(secLabel);
            var displayName = (fn + " " + ln).trim() || userIdStr;

            if (isNew) {
                state.users.push({
                    id: uid("u"),
                    userId: userIdStr,
                    name: displayName,
                    role: al === 1 ? "Administrator" : al === 3 ? "Read Only" : "Operator",
                    accessLevel: al,
                    authenticationMethod: authM,
                    password: authM === 1 ? "" : pw,
                    jobTitle: job,
                    userType: userTypeLabelFromAccessLevel(al),
                    firstName: fn,
                    middleInit: mi.slice(0, 1),
                    lastName: ln,
                    lastLogon: ""
                });
            } else {
                var tgt = state.users[editIndex];
                tgt.userId = userIdStr;
                tgt.firstName = fn;
                tgt.middleInit = mi.slice(0, 1);
                tgt.lastName = ln;
                tgt.jobTitle = job;
                tgt.name = displayName;
                tgt.role = al === 1 ? "Administrator" : al === 3 ? "Read Only" : "Operator";
                tgt.accessLevel = al;
                tgt.authenticationMethod = authM;
                tgt.userType = userTypeLabelFromAccessLevel(al);
                if (authM === 0) {
                    if (pw !== "") tgt.password = pw;
                } else {
                    tgt.password = "";
                }
            }
            saveState();
            closeStackedAppModal();
            var selAfter = isNew ? state.users.length - 1 : editIndex;
            refreshUserMaintenanceContent({ selectIndex: selAfter });
            toast("User saved.");
        });

        setTimeout(function () {
            var el = document.getElementById("usFirstName");
            if (el) el.focus();
        }, 0);
    }

    function openUserMaintenance(opts) {
        opts = opts || {};
        if (!isSessionLoggedIn()) {
            toast("You must be logged in to open User Maintenance.");
            return;
        }
        openAppModal(
            "User Maintenance - Binventory Workstation",
            buildUserMaintenanceBodyHtml(),
            "",
            "modal-user-maintenance modal-footer-hidden modal-win-toolwindow"
        );
        bindUserMaintenanceEvents(opts);
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
            case "operators-manual": {
                var manualUrl = "https://binmaster.com/amfile/file/download/file/1383/product/4198/";
                var a = document.createElement("a");
                a.href = manualUrl;
                a.target = "_blank";
                a.rel = "noopener noreferrer";
                a.click();
                break;
            }
            default:
                break;
        }
    }

    function ensureDbReadOnlyTutorialState() {
        if (dbReadOnlyTutorial) return dbReadOnlyTutorial;
        dbReadOnlyTutorial = {
            active: false,
            completed: false,
            stepIndex: 0,
            enteredStepIndex: -1,
            timerId: null,
            autoStepTimerId: null,
            pendingAdvanceStepIndex: -1,
            pendingAdvanceSinceMs: 0,
            steps: [
                {
                    title: "Step 1",
                    message: "Click Start.",
                    selector: "#winStartBtn",
                    advanceWhen: function () {
                        var startMenu = document.getElementById("winStartMenu");
                        return !!(startMenu && !startMenu.hidden);
                    }
                },
                {
                    title: "Step 2",
                    message: 'Type "cmd" in the search box.',
                    selector: "#winStartSearch",
                    highlightShape: "rect",
                    pointerSide: "left",
                    inputSelector: "#winStartSearch",
                    advanceWhen: function () {
                        var inp = document.getElementById("winStartSearch");
                        var panel = document.getElementById("winSearchPanelCmd");
                        var q = String(inp && inp.value ? inp.value : "").trim().toLowerCase();
                        return !!(panel && !panel.hidden && q.indexOf("cmd") === 0);
                    }
                },
                {
                    title: "Step 3",
                    message: 'Click Open (or press Enter) to launch Terminal.',
                    selector: "#winSearchPanelCmd [data-sim-cmd='open']",
                    highlightShape: "rect",
                    pointerSide: "left",
                    allowEnterSelector: "#winStartSearch",
                    advanceWhen: function () {
                        var shell = document.getElementById("appModalShell");
                        return !!(shell && shell.classList.contains("modal-command-prompt"));
                    }
                },
                {
                    title: "Step 4",
                    message: 'In Terminal, type "ipconfig".',
                    selector: "#cmdLine",
                    highlightShape: "rect",
                    pointerSide: "left",
                    inputSelector: "#cmdInput",
                    blockEnter: true,
                    advanceWhen: function () {
                        var inp = document.getElementById("cmdInput");
                        var q = String(inp && inp.value ? inp.value : "").trim().toLowerCase();
                        if (q === "ipconfig") return true;
                        var out = document.getElementById("cmdOutput");
                        var text = String(out && out.textContent ? out.textContent : "").toLowerCase();
                        return text.indexOf("h:\\>ipconfig") >= 0 || text.indexOf("\nip configuration") >= 0;
                    }
                },
                {
                    title: "Step 5",
                    message: "Press Enter.",
                    selector: "#cmdLine",
                    highlightShape: "rect",
                    pointerSide: "left",
                    allowEnterSelector: "#cmdInput",
                    advanceDelayMs: 240,
                    advanceWhen: function () {
                        var out = document.getElementById("cmdOutput");
                        return !!(out && String(out.textContent || "").indexOf("IPv4 Address") >= 0);
                    }
                },
                {
                    title: "Step 6",
                    message:
                        "This is the SERVER Computer IP. We will set this in Binventory. Click Next.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    requiresNext: true,
                    getTargetRect: function () {
                        var cmdScroll = document.getElementById("cmdScroll");
                        if (cmdScroll && cmdScroll.scrollTop > 2) return null;
                        if (this._readyAtMs && Date.now() < this._readyAtMs) return null;
                        return getCmdIpv4HighlightRect();
                    },
                    onEnter: function () {
                        var cmdScroll = document.getElementById("cmdScroll");
                        this._readyAtMs = Date.now() + 180;
                        if (!cmdScroll) return;
                        if (typeof cmdScroll.scrollTo === "function") {
                            cmdScroll.scrollTo({ top: 0, behavior: "smooth" });
                        } else {
                            cmdScroll.scrollTop = 0;
                        }
                    }
                },
                {
                    title: "Step 7",
                    message:
                        "Your IPv4 is " +
                        SIM_WORKSTATION_IPV4 +
                        ". Note it, then click X to close Terminal.",
                    selector: "#appModalShell.modal-command-prompt .app-modal-cap-close[data-close-app]",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var shell = document.getElementById("appModalShell");
                        return !(shell && shell.classList.contains("modal-command-prompt"));
                    }
                },
                {
                    title: "Step 8",
                    message: "Click Specifications.",
                    selector: "#menuBtnSpecifications",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var btn = document.getElementById("menuBtnSpecifications");
                        var root = btn && btn.closest ? btn.closest(".menu-root") : null;
                        return !!(root && root.classList.contains("open"));
                    }
                },
                {
                    title: "Step 9",
                    message: "Click Site Maintenance.",
                    selector: "#menuItemSiteMaintenance",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var shell = document.getElementById("appModalShell");
                        return !!(shell && shell.classList.contains("modal-site-maintenance"));
                    }
                },
                {
                    title: "Step 10",
                    message: "Select Current Location in the list.",
                    selector: ".sm-table tbody tr[data-sid='st1']",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var row = document.querySelector(".sm-table tbody tr[data-sid='st1']");
                        return !!(row && row.classList.contains("sm-selected"));
                    }
                },
                {
                    title: "Step 11",
                    message: "Click Select.",
                    selector: "#smBtnSelect",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var baseShell = document.getElementById("appModalShell");
                        var stackShell = document.getElementById("appModalShellStack");
                        var stackBackdrop = document.getElementById("backdropAppStack");
                        return !!(
                            (stackShell &&
                                stackShell.classList.contains("modal-site-setup") &&
                                stackBackdrop &&
                                stackBackdrop.classList.contains("show")) ||
                            (baseShell && baseShell.classList.contains("modal-site-setup"))
                        );
                    }
                },
                {
                    title: "Step 12",
                    message: "Clear the current Host IP Address field first.",
                    selector: "#ssuHostIp",
                    highlightShape: "rect",
                    pointerSide: "left",
                    inputSelector: "#ssuHostIp",
                    advanceWhen: function () {
                        var inp = document.getElementById("ssuHostIp");
                        return !!(inp && String(inp.value || "").trim() === "");
                    }
                },
                {
                    title: "Step 13",
                    message: "Type " + SIM_WORKSTATION_IPV4 + " for Host IP Address.",
                    selector: "#ssuHostIp",
                    highlightShape: "rect",
                    pointerSide: "left",
                    inputSelector: "#ssuHostIp",
                    advanceWhen: function () {
                        var inp = document.getElementById("ssuHostIp");
                        return !!(inp && String(inp.value || "").trim() === SIM_WORKSTATION_IPV4);
                    }
                },
                {
                    title: "Step 14",
                    message: "Click Save.",
                    selector: "#ssuSave",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var s = findCurrentSite();
                        var shell = document.getElementById("appModalShell");
                        return !!(
                            s &&
                            String(s.serviceHostIp || "").trim() === SIM_WORKSTATION_IPV4 &&
                            !(shell && shell.classList.contains("modal-site-setup"))
                        );
                    }
                },
                {
                    title: "Step 15",
                    message: "Click X to close Site Maintenance.",
                    selector: "#appModalShell.modal-site-maintenance .app-modal-cap-close[data-close-app]",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var shell = document.getElementById("appModalShell");
                        return !(shell && shell.classList.contains("modal-site-maintenance"));
                    }
                },
                {
                    title: "Step 16",
                    message: "Click X to close Binventory.",
                    selector: ".title-btns button[data-sim='close']",
                    highlightShape: "rect",
                    pointerSide: "left",
                    /* advanceWhen used to be “desktop mode only”, which could match without a real Exit; require exitToSimDesktop(). */
                    dbRoRequireRecordedExit: true,
                    onEnter: function () {
                        state._dbRoTutorialClosedAppForStep16 = false;
                    },
                    advanceWhen: function () {
                        var wrap = document.getElementById("pageWrap");
                        return !!(
                            wrap &&
                            wrap.classList.contains("page-wrap--desktop-mode") &&
                            state._dbRoTutorialClosedAppForStep16
                        );
                    }
                },
                {
                    title: "Step 17",
                    message: "Double-click the eBob desktop icon to reopen.",
                    selector: "#desktopEbobIcon",
                    highlightShape: "rect",
                    pointerSide: "left",
                    onEnter: function () {
                        var wrap = document.getElementById("pageWrap");
                        state._dbRoStep17BeganOnDesktop = !!(
                            wrap && wrap.classList.contains("page-wrap--desktop-mode")
                        );
                    },
                    advanceWhen: function () {
                        var wrap = document.getElementById("pageWrap");
                        var isDesktop = !!(wrap && wrap.classList.contains("page-wrap--desktop-mode"));
                        if (isDesktop) return false;
                        if (!state._dbRoStep17BeganOnDesktop) return false;
                        return !document.querySelector(".vessel-status-readonly");
                    }
                },
                {
                    title: "Step 18",
                    message: "Click Vessel.",
                    selector: "#menuBtnVessel",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        var btn = document.getElementById("menuBtnVessel");
                        var root = btn && btn.closest ? btn.closest(".menu-root") : null;
                        return !!(root && root.classList.contains("open"));
                    }
                },
                {
                    title: "Step 19",
                    message: "Click Measure All Vessels.",
                    selector: "#menuItemMeasureAll",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        return areAnyVesselsMeasuring();
                    }
                },
                {
                    title: "Step 20",
                    message:
                        'Status should show "Retracted". This means the reading was successful after the fix. Click Finish',
                    highlightShape: "rect",
                    pointerSide: "left",
                    requiresNext: true,
                    nextLabel: "Finish",
                    getTargetRect: function () {
                        if (areAnyVesselsMeasuring()) return null;
                        return getRetractedStatusRect();
                    },
                    onNext: function () {
                        if (areAnyVesselsMeasuring()) {
                            toast("Waiting for Measure All to complete...");
                            return false;
                        }
                        if (!getRetractedStatusRect()) {
                            toast('Waiting for status to show "Retracted"...');
                            return false;
                        }
                        finishDbReadOnlyTutorial();
                        return true;
                    }
                }
            ]
        };
        return dbReadOnlyTutorial;
    }

    function getSvcMscBackdropEl() {
        return document.getElementById("backdropServicesMsc");
    }

    function isSvcMscVisible() {
        var bd = getSvcMscBackdropEl();
        return !!(bd && bd.classList.contains("show"));
    }

    /** Center eBob rows in the Services list and prevent scrolling (wheel / touch / scroll offset). */
    function installPendingUnknownServicesScrollLock() {
        var wrap = document.getElementById("svcMscTableScrollWrap");
        if (!wrap) return;
        if (wrap.__puScrollLockActive) {
            pendingUnknownSnapServicesScrollToEbob();
            return;
        }
        wrap.__puScrollLockActive = true;

        function lockToIdeal() {
            if (wrap.__puLockIdeal == null) return;
            wrap.scrollTop = wrap.__puLockIdeal;
        }

        function onWheelTouch(e) {
            e.preventDefault();
            lockToIdeal();
        }

        function onScroll() {
            lockToIdeal();
        }

        wrap.addEventListener("scroll", onScroll, { passive: true });
        wrap.addEventListener("wheel", onWheelTouch, { passive: false });
        wrap.addEventListener("touchmove", onWheelTouch, { passive: false });

        wrap.__puScrollLockOnScroll = onScroll;
        wrap.__puScrollLockOnWheelTouch = onWheelTouch;

        pendingUnknownSnapServicesScrollToEbob();
        window.requestAnimationFrame(function () {
            pendingUnknownSnapServicesScrollToEbob();
        });
    }

    function pendingUnknownSnapServicesScrollToEbob() {
        var wrap = document.getElementById("svcMscTableScrollWrap");
        if (!wrap) return;
        var row = wrap.querySelector('tbody tr[data-service-id="ebob-engine"]');
        if (!row) return;
        var target = Math.max(
            0,
            row.offsetTop - wrap.clientHeight / 2 + row.offsetHeight / 2
        );
        wrap.scrollTop = target;
        wrap.__puLockIdeal = target;
    }

    function clearPendingUnknownServicesScrollLock() {
        var wrap = document.getElementById("svcMscTableScrollWrap");
        if (!wrap) return;
        if (wrap.__puScrollLockOnScroll) {
            wrap.removeEventListener("scroll", wrap.__puScrollLockOnScroll);
        }
        if (wrap.__puScrollLockOnWheelTouch) {
            wrap.removeEventListener("wheel", wrap.__puScrollLockOnWheelTouch);
            wrap.removeEventListener("touchmove", wrap.__puScrollLockOnWheelTouch);
        }
        wrap.__puScrollLockActive = false;
        wrap.__puLockIdeal = null;
        wrap.__puScrollLockOnScroll = null;
        wrap.__puScrollLockOnWheelTouch = null;
    }

    /** Name column text only (for “locate Engine + Scheduler” step; avoids spanning into Description column / off-screen). */
    function getSvcMscServiceRowNameOnlyRawRect(serviceId) {
        var tr = document.querySelector(
            '#backdropServicesMsc tr[data-service-id="' + serviceId + '"]'
        );
        if (!tr || !isVisibleForTutorial(tr)) return null;
        var nameEl = tr.querySelector("td:nth-child(1) .svc-msc-name-text");
        if (nameEl && isVisibleForTutorial(nameEl)) {
            var r = nameEl.getBoundingClientRect();
            return { left: r.left, top: r.top, width: r.width, height: r.height };
        }
        var td = tr.querySelector("td:nth-child(1)");
        if (td && isVisibleForTutorial(td)) {
            var r2 = td.getBoundingClientRect();
            return { left: r2.left, top: r2.top, width: r2.width, height: r2.height };
        }
        return null;
    }

    /** Padded highlight for right-click / row steps: Name column only (same tight box as locate step). */
    function getSvcMscServiceRowHighlightRect(serviceId) {
        var r = getSvcMscServiceRowNameOnlyRawRect(serviceId);
        if (!r) return null;
        var pad = 4;
        return {
            left: r.left - pad,
            top: r.top - pad,
            width: r.width + pad * 2,
            height: r.height + pad * 2
        };
    }

    /** Single outer box around eBob Engine + eBob Scheduler **name** cells only (not Description column). */
    function getSvcMscEbobEngineSchedulerPairHighlightRect() {
        var ra = getSvcMscServiceRowNameOnlyRawRect("ebob-engine");
        var rb = getSvcMscServiceRowNameOnlyRawRect("ebob-scheduler");
        if (!ra || !rb) return null;
        var left = Math.min(ra.left, rb.left);
        var top = Math.min(ra.top, rb.top);
        var right = Math.max(ra.left + ra.width, rb.left + rb.width);
        var bottom = Math.max(ra.top + ra.height, rb.top + rb.height);
        var pad = 4;
        return {
            left: left - pad,
            top: top - pad,
            width: right - left + pad * 2,
            height: bottom - top + pad * 2
        };
    }

    function getVesselStatusTextNormalized(vid) {
        var v = findVessel(vid);
        return v && v.status ? String(v.status).trim().toLowerCase() : "";
    }

    function ensurePendingUnknownTutorialState() {
        if (pendingUnknownTutorial) return pendingUnknownTutorial;
        pendingUnknownTutorial = {
            active: false,
            completed: false,
            stepIndex: 0,
            enteredStepIndex: -1,
            timerId: null,
            autoStepTimerId: null,
            pendingAdvanceStepIndex: -1,
            pendingAdvanceSinceMs: 0,
            steps: [
                {
                    title: "Step 1a",
                    message:
                        'Click Measure on the silo that shows "Unknown".',
                    selector: '.vessel[data-vessel-id="v1"] [data-vessel-action="measure"]',
                    highlightShape: "rect",
                    pointerSide: "left",
                    /* Session restore can leave _puGlitchDemoDone true while the tutorial restarts at 1a — clear on entry. */
                    onEnter: function () {
                        state._puGlitchDemoDone = false;
                    },
                    advanceWhen: function () {
                        return !!state._puGlitchDemoDone;
                    },
                    advanceDelayMs: 400
                },
                {
                    title: "Step 1b",
                    message:
                        'Notice: status flashed Pending, then returned to Unknown. Click Next.',
                    selector: '.vessel[data-vessel-id="v1"] .vessel-status-text',
                    highlightShape: "rect",
                    pointerSide: "left",
                    requiresNext: true
                },
                {
                    title: "Step 2",
                    message:
                        'Notice how this silo is Pending? We will address this as well. Both issues have the same resolution. Click Next.',
                    selector: '.vessel[data-vessel-id="v2"] .vessel-status-text',
                    highlightShape: "rect",
                    pointerSide: "left",
                    requiresNext: true
                },
                {
                    title: "Step 3",
                    message: "Click Start.",
                    selector: "#winStartBtn",
                    advanceWhen: function () {
                        var sm = document.getElementById("winStartMenu");
                        return !!(sm && !sm.hidden);
                    }
                },
                {
                    title: "Step 4",
                    message: 'Type "services" in the search box.',
                    selector: "#winStartSearch",
                    highlightShape: "rect",
                    pointerSide: "left",
                    inputSelector: "#winStartSearch",
                    advanceWhen: function () {
                        var inp = document.getElementById("winStartSearch");
                        var panel = document.getElementById("winSearchPanelServices");
                        var q = String(inp && inp.value ? inp.value : "").trim().toLowerCase();
                        return !!(panel && !panel.hidden && (q === "services" || q.indexOf("services") === 0));
                    }
                },
                {
                    title: "Step 5",
                    message: 'Click "Run as administrator" (do not use Open — elevated Services is required).',
                    selector: '#winSearchPanelServices [data-sim-svc="admin"]',
                    highlightShape: "rect",
                    pointerSide: "left",
                    blockSimSvcOpen: true,
                    advanceWhen: function () {
                        var uac = document.getElementById("backdropUac");
                        return !!(uac && uac.classList.contains("show"));
                    }
                },
                {
                    title: "Step 6",
                    message:
                        "User Account Control: click Yes. On some systems, admin rights may require signing in.",
                    selector: "#uacYes",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        return isSvcMscVisible();
                    }
                },
                {
                    title: "Step 7",
                    message:
                        "Locate eBob Engine and eBob Scheduler Services. When both are highlighted in the box, click Next.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#ebobTutorialNext",
                    getTargetRect: function () {
                        return getSvcMscEbobEngineSchedulerPairHighlightRect();
                    },
                    requiresNext: true,
                    blockSvcMscTableRows: true,
                    restrictSvcMscToCtxMenuFlow: true,
                    onEnter: function () {
                        installPendingUnknownServicesScrollLock();
                    }
                },
                {
                    title: "Step 8",
                    message:
                        "Right-click the eBob Engine Service row.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: '#backdropServicesMsc tr[data-service-id="ebob-engine"]',
                    getTargetRect: function () {
                        return getSvcMscServiceRowHighlightRect("ebob-engine");
                    },
                    blockSvcMscTableRows: true,
                    svcCtxRowIdForRightClick: "ebob-engine",
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        var menu = document.getElementById("svcCtxMenu");
                        return !!(menu && !menu.hidden);
                    }
                },
                {
                    title: "Step 9",
                    message: "Click Stop.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#svcCtxStop",
                    getTargetRect: function () {
                        var el = document.getElementById("svcCtxStop");
                        if (isVisibleForTutorial(el)) return el;
                        return getSvcMscServiceRowHighlightRect("ebob-engine");
                    },
                    blockSvcMscTableRows: true,
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        return !state.ebobServicesRunning;
                    }
                },
                {
                    title: "Step 10",
                    message: "Right-click eBob Engine Service again.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: '#backdropServicesMsc tr[data-service-id="ebob-engine"]',
                    getTargetRect: function () {
                        return getSvcMscServiceRowHighlightRect("ebob-engine");
                    },
                    blockSvcMscTableRows: true,
                    svcCtxRowIdForRightClick: "ebob-engine",
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        var menu = document.getElementById("svcCtxMenu");
                        return !!(menu && !menu.hidden);
                    }
                },
                {
                    title: "Step 11",
                    message: "Click Start.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#svcCtxStart",
                    getTargetRect: function () {
                        var el = document.getElementById("svcCtxStart");
                        if (isVisibleForTutorial(el)) return el;
                        return getSvcMscServiceRowHighlightRect("ebob-engine");
                    },
                    blockSvcMscTableRows: true,
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        return !!state.ebobServicesRunning;
                    }
                },
                {
                    title: "Step 12",
                    message: "Right-click eBob Scheduler Service.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: '#backdropServicesMsc tr[data-service-id="ebob-scheduler"]',
                    getTargetRect: function () {
                        return getSvcMscServiceRowHighlightRect("ebob-scheduler");
                    },
                    blockSvcMscTableRows: true,
                    svcCtxRowIdForRightClick: "ebob-scheduler",
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        var menu = document.getElementById("svcCtxMenu");
                        return !!(menu && !menu.hidden);
                    }
                },
                {
                    title: "Step 13",
                    message: "Click Start.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#svcCtxStart",
                    getTargetRect: function () {
                        var el = document.getElementById("svcCtxStart");
                        if (isVisibleForTutorial(el)) return el;
                        return getSvcMscServiceRowHighlightRect("ebob-scheduler");
                    },
                    blockSvcMscTableRows: true,
                    restrictSvcMscToCtxMenuFlow: true,
                    advanceWhen: function () {
                        return !!(state.ebobServicesRunning && state.ebobSchedulerRunning);
                    }
                },
                {
                    title: "Step 14",
                    message: "Close the Services window (X).",
                    selector: "#svcMscClose",
                    highlightShape: "rect",
                    pointerSide: "left",
                    advanceWhen: function () {
                        return !isSvcMscVisible();
                    },
                    advanceDelayMs: 220
                },
                {
                    title: "Step 15",
                    message:
                        'Database Read Only displays after restarting eBob Services. Click Close and Relaunch eBob. Click Next to continue.',
                    hideTutorialHighlight: true,
                    requiresNext: true,
                    onNext: function () {
                        state._puAwaitingPostRestartReadOnlyAck = true;
                        return true;
                    }
                },
                {
                    title: "Step 16",
                    message: "Close Binventory — click X on the title bar, then you will relaunch from the desktop.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: ".title-btns button[data-sim='close']",
                    getTargetRect: function () {
                        var wrap = document.getElementById("pageWrap");
                        if (wrap && wrap.classList.contains("page-wrap--desktop-mode")) return null;
                        var closeBtn = document.querySelector(".title-btns button[data-sim='close']");
                        return isVisibleForTutorial(closeBtn) ? closeBtn : null;
                    },
                    onEnter: function () {
                        clearPendingUnknownServicesScrollLock();
                    },
                    advanceWhen: function () {
                        var wrap = document.getElementById("pageWrap");
                        return !!(wrap && wrap.classList.contains("page-wrap--desktop-mode"));
                    },
                    advanceDelayMs: 160
                },
                {
                    title: "Step 17",
                    message: "Double-click eBob Icon to Launch.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#desktopEbobIcon",
                    getTargetRect: function () {
                        var wrap = document.getElementById("pageWrap");
                        if (!wrap || !wrap.classList.contains("page-wrap--desktop-mode")) return null;
                        var icon = document.querySelector("#desktopEbobIcon");
                        return isVisibleForTutorial(icon) ? icon : null;
                    },
                    advanceWhen: function () {
                        var wrap = document.getElementById("pageWrap");
                        return !!(wrap && !wrap.classList.contains("page-wrap--desktop-mode"));
                    },
                    advanceDelayMs: 220
                },
                {
                    title: "Step 18",
                    message:
                        'After the workstation loads, click Measure on the vessel that used to show "Unknown".',
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: '.vessel[data-vessel-id="v1"] [data-vessel-action="measure"]',
                    onEnter: function () {
                        if (!state._puPostRestartResolved) {
                            state._puAwaitingPostRestartReadOnlyAck = false;
                            applyPendingUnknownPostRestartState();
                        }
                        var v = findVessel("v1");
                        state._puStep18BaselineLastMeasurement = v && v.lastMeasurement ? String(v.lastMeasurement) : "";
                    },
                    advanceWhen: function () {
                        var v = findVessel("v1");
                        if (!v) return false;
                        var base = state._puStep18BaselineLastMeasurement || "";
                        var cur = v.lastMeasurement ? String(v.lastMeasurement) : "";
                        return cur !== base && String(v.status || "").trim().toLowerCase() === "ready";
                    }
                },
                {
                    title: "Step 19",
                    message:
                        'Then click Measure on the silo that showed Pending.',
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: '.vessel[data-vessel-id="v2"] [data-vessel-action="measure"]',
                    onEnter: function () {
                        var v = findVessel("v2");
                        if (v && state._puPostRestartResolved) {
                            delete v.tutorialPuGlitch;
                            delete v.tutorialPuStuck;
                            var sid = parseInt(v.sensorTypeId, 10) || 11;
                            v.status = statusStringForDashboard(sid, 55);
                            v.name = "NCR-80 Unknown";
                            saveState();
                            refreshUI();
                        }
                        state._puStep19BaselineLastMeasurement = v && v.lastMeasurement ? String(v.lastMeasurement) : "";
                    },
                    advanceWhen: function () {
                        var v = findVessel("v2");
                        if (!v) return false;
                        var base = state._puStep19BaselineLastMeasurement || "";
                        var cur = v.lastMeasurement ? String(v.lastMeasurement) : "";
                        return cur !== base && String(v.status || "").trim().toLowerCase() === "ready";
                    }
                },
                {
                    title: "Step 20",
                    message: "Open Vessel, then choose Measure All Vessels.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#menuBtnVessel",
                    advanceWhen: function () {
                        var btn = document.getElementById("menuBtnVessel");
                        var root = btn && btn.closest ? btn.closest(".menu-root") : null;
                        return !!(root && root.classList.contains("open"));
                    }
                },
                {
                    title: "Step 21",
                    message: "Click Measure All Vessels.",
                    highlightShape: "rect",
                    pointerSide: "left",
                    selector: "#menuItemMeasureAll",
                    onEnter: function () {
                        state._puTutorialMeasureAllStarted = false;
                    },
                    advanceWhen: function () {
                        if (areAnyVesselsMeasuring()) {
                            state._puTutorialMeasureAllStarted = true;
                            return false;
                        }
                        return !!(
                            state._puTutorialMeasureAllStarted &&
                            !areAnyVesselsMeasuring() &&
                            pendingUnknownAllMeasurableSiteVesselsReady()
                        );
                    },
                    advanceDelayMs: 320
                },
                {
                    title: "Step 22",
                    message:
                        "All vessels are now measuring properly. Click Finish to complete troubleshooting guide.",
                    hideTutorialHighlight: true,
                    requiresNext: true,
                    nextLabel: "Finish",
                    onNext: function () {
                        if (areAnyVesselsMeasuring()) {
                            toast("Waiting for Measure All to complete...");
                            return false;
                        }
                        if (!pendingUnknownAllMeasurableSiteVesselsReady()) {
                            toast("Waiting for all silos to show Ready...");
                            return false;
                        }
                        finishDbReadOnlyTutorial();
                        return true;
                    }
                }
            ]
        };
        return pendingUnknownTutorial;
    }

    function guidedGetActiveTutorialState() {
        if (activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY) return ensureDbReadOnlyTutorialState();
        if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) return ensurePendingUnknownTutorialState();
        return null;
    }

    function getDbTutorialDom() {
        return {
            tooltip: document.getElementById("ebobTutorialTooltip"),
            step: document.getElementById("ebobTutorialStep"),
            title: document.getElementById("ebobTutorialTitle"),
            msg: document.getElementById("ebobTutorialMsg"),
            next: document.getElementById("ebobTutorialNext"),
            pointer: document.getElementById("ebobTutorialPointer"),
            circle: document.getElementById("ebobTutorialCircle"),
            complete: document.getElementById("ebobTutorialComplete"),
            completeDone: document.getElementById("ebobTutorialCompleteDone")
        };
    }

    function setDbTutorialVisibility(show) {
        var dom = getDbTutorialDom();
        [dom.tooltip, dom.pointer, dom.circle].forEach(function (el) {
            if (!el) return;
            if (show) el.classList.remove("ebob-tutorial-hidden");
            else el.classList.add("ebob-tutorial-hidden");
        });
        if (dom.next && !show) {
            dom.next.classList.remove("ebob-tutorial-next--flash");
        }
    }

    function isVisibleForTutorial(el) {
        if (!el) return false;
        if (el.hidden) return false;
        var cs = window.getComputedStyle(el);
        if (cs.display === "none" || cs.visibility === "hidden") return false;
        var r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
    }

    function getDbTutorialTarget(step) {
        if (!step) return null;
        if (typeof step.getTargetRect === "function") {
            return step.getTargetRect() || null;
        }
        if (!step.selector) return null;
        var el = document.querySelector(step.selector);
        if (isVisibleForTutorial(el)) return el;
        return null;
    }

    function isGuidedTutorialLockActive() {
        var t = guidedGetActiveTutorialState();
        return !!(t && t.active && !t.completed);
    }

    function isTargetInsideSelector(target, selector) {
        if (!target || !selector) return false;
        if (!target.closest) return false;
        try {
            return !!target.closest(selector);
        } catch (e) {
            return false;
        }
    }

    function isDbTutorialInteractionAllowed(target, step, ev) {
        if (!step) return true;
        var evType = ev && ev.type ? ev.type : "";
        if (step.blockSimSvcOpen && isTargetInsideSelector(target, '[data-sim-svc="open"]')) return false;
        if (isTargetInsideSelector(target, "#ebobTutorialTooltip")) return true;
        if (isTargetInsideSelector(target, "#ebobTutorialNext")) return true;
        if (isTargetInsideSelector(target, "#guideBackToMenu")) return true;

        var svcDataRow =
            target && target.closest ? target.closest("#backdropServicesMsc tbody tr[data-service-id]") : null;
        if (step.blockSvcMscTableRows && svcDataRow) {
            if (evType === "contextmenu") {
                var rid = step.svcCtxRowIdForRightClick;
                if (!rid) return false;
                return svcDataRow.getAttribute("data-service-id") === rid;
            }
            if (evType === "pointerdown" || evType === "auxclick") {
                if (ev && ev.button !== 0) {
                    var ridBtn = step.svcCtxRowIdForRightClick;
                    return !!(ridBtn && svcDataRow.getAttribute("data-service-id") === ridBtn);
                }
                return false;
            }
            if (evType === "click" || evType === "dblclick") return false;
        }

        if (step.selector && isTargetInsideSelector(target, step.selector)) return true;
        if (step.inputSelector && isTargetInsideSelector(target, step.inputSelector)) return true;
        if (step.allowEnterSelector && isTargetInsideSelector(target, step.allowEnterSelector)) return true;

        if (step.restrictSvcMscToCtxMenuFlow && isTargetInsideSelector(target, "#backdropServicesMsc")) {
            if (isTargetInsideSelector(target, "#svcMscTableScrollWrap") && !svcDataRow) return true;
            return false;
        }

        return false;
    }

    /**
     * Services sim dismisses the context menu on any document capture-phase click outside #svcCtxMenu.
     * That includes tutorial chrome and the services table — users can spam-click and lose the menu,
     * so #svcCtxStop/#svcCtxStart become invisible and the guide pointer disappears. Skip auto-dismiss
     * while the guided step requires those menu actions.
     */
    function shouldSuppressSvcCtxMenuDismissForTutorial(e) {
        if (!isGuidedTutorialLockActive()) return false;
        if (e && e.target && e.target.closest) {
            if (e.target.closest("#ebobTutorialTooltip")) return true;
            if (e.target.closest("#ebobTutorialPointer")) return true;
            if (e.target.closest("#ebobTutorialCircle")) return true;
        }
        var gt = guidedGetActiveTutorialState();
        var st = gt && gt.steps[gt.stepIndex];
        if (!st || !st.restrictSvcMscToCtxMenuFlow) return false;
        var sel = st.selector != null ? String(st.selector) : "";
        return sel === "#svcCtxStop" || sel === "#svcCtxStart";
    }

    function bindDbReadOnlyTutorialGuards() {
        if (dbTutorialGuardsBound) return;
        dbTutorialGuardsBound = true;

        function blockIfLocked(e) {
            if (!isGuidedTutorialLockActive()) return;
            var t = guidedGetActiveTutorialState();
            var step = t.steps[t.stepIndex];
            if (!step) return;
            if (isDbTutorialInteractionAllowed(e.target, step, e)) return;
            e.preventDefault();
            e.stopPropagation();
            if (e.stopImmediatePropagation) e.stopImmediatePropagation();
        }

        document.addEventListener("pointerdown", blockIfLocked, true);
        document.addEventListener("click", blockIfLocked, true);
        document.addEventListener("contextmenu", blockIfLocked, true);
        document.addEventListener("dblclick", blockIfLocked, true);
        document.addEventListener("auxclick", blockIfLocked, true);

        document.addEventListener(
            "keydown",
            function (e) {
                if (!isGuidedTutorialLockActive()) return;
                var t = guidedGetActiveTutorialState();
                var step = t.steps[t.stepIndex];
                if (!step) return;
                if (e.key === "Escape") {
                    e.preventDefault();
                    e.stopPropagation();
                    if (e.stopImmediatePropagation) e.stopImmediatePropagation();
                    return;
                }
                var activeEl = document.activeElement;
                if (step.blockEnter && e.key === "Enter") {
                    if (!isTargetInsideSelector(activeEl, step.allowEnterSelector || "")) {
                        e.preventDefault();
                        e.stopPropagation();
                        if (e.stopImmediatePropagation) e.stopImmediatePropagation();
                        return;
                    }
                    if (step.inputSelector && isTargetInsideSelector(activeEl, step.inputSelector)) {
                        e.preventDefault();
                        e.stopPropagation();
                        if (e.stopImmediatePropagation) e.stopImmediatePropagation();
                        return;
                    }
                }
                if (step.allowEnterSelector && e.key === "Enter") {
                    if (!isTargetInsideSelector(activeEl, step.allowEnterSelector)) {
                        e.preventDefault();
                        e.stopPropagation();
                        if (e.stopImmediatePropagation) e.stopImmediatePropagation();
                    }
                }
            },
            true
        );
    }

    function positionDbTutorialPointer(target) {
        var dom = getDbTutorialDom();
        var gt = guidedGetActiveTutorialState();
        if (!gt || !gt.active || gt.completed) {
            if (dom.pointer) dom.pointer.classList.add("ebob-tutorial-hidden");
            if (dom.circle) dom.circle.classList.add("ebob-tutorial-hidden");
            return;
        }
        var stepNoHighlight = gt.steps[gt.stepIndex] || {};
        if (stepNoHighlight.hideTutorialHighlight) {
            if (dom.pointer) dom.pointer.classList.add("ebob-tutorial-hidden");
            if (dom.circle) dom.circle.classList.add("ebob-tutorial-hidden");
            return;
        }
        if (!target || !dom.pointer || !dom.circle) {
            if (dom.pointer) dom.pointer.classList.add("ebob-tutorial-hidden");
            if (dom.circle) dom.circle.classList.add("ebob-tutorial-hidden");
            return;
        }
        var t = gt;
        var step = t.steps[t.stepIndex] || {};
        var highlightShape = step.highlightShape === "rect" ? "rect" : "circle";
        var rect = null;
        var isRectObject =
            target &&
            typeof target.left === "number" &&
            typeof target.top === "number" &&
            typeof target.width === "number" &&
            typeof target.height === "number";
        if (isRectObject) {
            rect = {
                left: target.left,
                top: target.top,
                width: target.width,
                height: target.height
            };
        } else {
            rect = target.getBoundingClientRect();
        }
        var cx = rect.left + rect.width / 2;
        var cy = rect.top + rect.height / 2;
        dom.circle.classList.remove("is-circle", "is-rect");
        var pointerSide = String(step.pointerSide || "auto").toLowerCase();

        var pointerW = dom.pointer.offsetWidth || 48;
        if (highlightShape === "rect") {
            var pad = 6;
            var hLeft = Math.max(4, rect.left - pad);
            var hTop = Math.max(4, rect.top - pad);
            var hW = Math.min(window.innerWidth - hLeft - 4, rect.width + pad * 2);
            var hH = Math.min(window.innerHeight - hTop - 4, rect.height + pad * 2);
            dom.circle.classList.add("is-rect");
            dom.circle.style.left = hLeft + "px";
            dom.circle.style.top = hTop + "px";
            dom.circle.style.width = hW + "px";
            dom.circle.style.height = hH + "px";

            var rectLeftX = hLeft - pointerW - 10;
            var rectRightX = hLeft + hW + 10;
            var canLeft = rectLeftX >= 4;
            var canRight = rectRightX + pointerW <= window.innerWidth - 4;
            var useLeft =
                pointerSide === "left"
                    ? canLeft
                    : pointerSide === "right"
                      ? !canRight
                      : canLeft;
            if (!canLeft && canRight) useLeft = false;
            if (!canRight && canLeft) useLeft = true;
            if (useLeft) {
                dom.pointer.textContent = "👉";
                dom.pointer.style.left = rectLeftX + "px";
            } else {
                dom.pointer.textContent = "👈";
                dom.pointer.style.left = rectRightX + "px";
            }
            dom.pointer.style.top = hTop + Math.max(0, hH / 2 - 24) + "px";
        } else {
            var radius = 28;
            dom.circle.classList.add("is-circle");
            dom.circle.style.width = "56px";
            dom.circle.style.height = "56px";
            dom.circle.style.left = cx - radius + "px";
            dom.circle.style.top = cy - radius + "px";
            var leftX = cx - radius - pointerW - 8;
            var rightX = cx + radius + 8;
            if (leftX >= 4) {
                dom.pointer.textContent = "👉";
                dom.pointer.style.left = leftX + "px";
            } else {
                dom.pointer.textContent = "👈";
                dom.pointer.style.left = rightX + "px";
            }
            dom.pointer.style.top = cy - 24 + "px";
        }
        dom.pointer.classList.remove("ebob-tutorial-hidden");
        dom.circle.classList.remove("ebob-tutorial-hidden");
    }

    function renderDbReadOnlyTutorial() {
        var t = guidedGetActiveTutorialState();
        if (!t || !t.active || t.completed) return;
        var dom = getDbTutorialDom();
        if (!dom.tooltip || !dom.step || !dom.title || !dom.msg) return;
        var step = t.steps[t.stepIndex];
        if (!step) return;
        if (t.enteredStepIndex !== t.stepIndex) {
            t.enteredStepIndex = t.stepIndex;
            if (typeof step.onEnter === "function") {
                step.onEnter();
            }
        }
        /* Match progress line to the heading (e.g. "Step 18 of 23" vs title "Step 18"), not stepIndex+1. */
        var rawTitle = step.title != null ? String(step.title).trim() : "";
        var stepLabel = /^Step\b/i.test(rawTitle) ? rawTitle : "Step " + (t.stepIndex + 1);
        dom.step.textContent = stepLabel + " of " + t.steps.length;
        dom.title.textContent = step.title;
        dom.msg.textContent = step.message;
        if (dom.next) {
            dom.next.hidden = !step.requiresNext;
            dom.next.textContent = step.nextLabel || "Next";
            /* Pulse Next/Finish whenever the user must click it (read-and-acknowledge steps). */
            var pulseNext = !!(step.requiresNext && step.suppressNextFlash !== true);
            if (pulseNext) {
                dom.next.classList.add("ebob-tutorial-next--flash");
            } else {
                dom.next.classList.remove("ebob-tutorial-next--flash");
            }
        }
        dom.tooltip.classList.remove("ebob-tutorial-hidden");
        positionDbTutorialPointer(getDbTutorialTarget(step));
    }

    function areAnyVesselsMeasuring() {
        return !!(state.vessels || []).some(function (v) {
            return v && v._measureTimers && v._measureTimers.length > 0;
        });
    }

    function getRetractedStatusRect() {
        var nodes = document.querySelectorAll(".vessel-status-val");
        for (var i = 0; i < nodes.length; i++) {
            var wrap = nodes[i];
            if (!isVisibleForTutorial(wrap)) continue;
            var inner = wrap.querySelector(".vessel-status-text");
            var el = inner || wrap;
            var txt = String(el.textContent || "").trim().toLowerCase();
            if (txt !== "retracted") continue;
            var r = el.getBoundingClientRect();
            return {
                left: r.left - 4,
                top: r.top - 4,
                width: r.width + 8,
                height: r.height + 8
            };
        }
        return null;
    }

    function showDbTutorialCompleteOverlay() {
        var dom = getDbTutorialDom();
        if (!dom.complete) return;
        var titleEl = dom.complete.querySelector(".ebob-tutorial-complete__title");
        var subEl = dom.complete.querySelector(".ebob-tutorial-complete__sub");
        if (titleEl && subEl) {
            if (activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN) {
                titleEl.textContent = "Status Restored";
                subEl.textContent =
                    "Pending and Unknown issues are resolved after restarting eBob services. You can return to the menu to replay or try another scenario.";
            } else {
                titleEl.textContent = "Binventory Error Resolved";
                subEl.textContent =
                    "Database Read Only solved and measurements are verified. Great work. You can now return to the menu to replay or solve a different issue.";
            }
        }
        dom.complete.classList.add("ebob-tutorial-complete--visible");
        dom.complete.setAttribute("aria-hidden", "false");
    }

    function hideDbTutorialCompleteOverlay() {
        var dom = getDbTutorialDom();
        if (!dom.complete) return;
        dom.complete.classList.remove("ebob-tutorial-complete--visible");
        dom.complete.setAttribute("aria-hidden", "true");
    }

    function returnToTutorialScenarioMenu() {
        try {
            sessionStorage.removeItem(EBOB_TUTORIAL_MODE_KEY);
            sessionStorage.removeItem(EBOB_TUTORIAL_STATE_KEY);
        } catch (e) {
            /* private mode */
        }
        var base = window.location.pathname || "ebob.html";
        window.location.href = base;
    }

    function finishDbReadOnlyTutorial() {
        var t = guidedGetActiveTutorialState();
        if (!t) return;
        t.completed = true;
        t.active = false;
        if (t.timerId != null) {
            clearInterval(t.timerId);
            t.timerId = null;
        }
        if (t.autoStepTimerId != null) {
            clearTimeout(t.autoStepTimerId);
            t.autoStepTimerId = null;
        }
        setDbTutorialVisibility(false);
        showDbTutorialCompleteOverlay();
    }

    function advanceDbReadOnlyTutorial() {
        var t = guidedGetActiveTutorialState();
        if (!t || !t.active || t.completed) return;
        if (t.autoStepTimerId != null) {
            clearTimeout(t.autoStepTimerId);
            t.autoStepTimerId = null;
        }
        t.pendingAdvanceStepIndex = -1;
        t.pendingAdvanceSinceMs = 0;
        t.stepIndex += 1;
        if (t.stepIndex >= t.steps.length) {
            t.completed = true;
            t.active = false;
            setDbTutorialVisibility(false);
            if (
                activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY ||
                activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN
            ) {
                showDbTutorialCompleteOverlay();
            }
            return;
        }
        renderDbReadOnlyTutorial();
    }

    function tickDbReadOnlyTutorial() {
        var t = guidedGetActiveTutorialState();
        if (!t || !t.active || t.completed) return;
        var step = t.steps[t.stepIndex];
        if (!step) return;
        renderDbReadOnlyTutorial();
        if (typeof step.advanceWhen === "function" && step.advanceWhen()) {
            var delayMs = Math.max(0, parseInt(step.advanceDelayMs, 10) || 0);
            if (delayMs > 0) {
                if (t.pendingAdvanceStepIndex !== t.stepIndex) {
                    t.pendingAdvanceStepIndex = t.stepIndex;
                    t.pendingAdvanceSinceMs = Date.now();
                    return;
                }
                if (Date.now() - t.pendingAdvanceSinceMs < delayMs) {
                    return;
                }
            }
            t.pendingAdvanceStepIndex = -1;
            t.pendingAdvanceSinceMs = 0;
            advanceDbReadOnlyTutorial();
            return;
        } else {
            t.pendingAdvanceStepIndex = -1;
            t.pendingAdvanceSinceMs = 0;
        }
        if (step.autoAdvanceMs && t.autoStepTimerId == null) {
            t.autoStepTimerId = setTimeout(function () {
                t.autoStepTimerId = null;
                if (!t.active || t.completed) return;
                if (t.stepIndex >= t.steps.length) return;
                if (t.steps[t.stepIndex] === step) {
                    advanceDbReadOnlyTutorial();
                }
            }, step.autoAdvanceMs);
        }
    }

    function stopDbReadOnlyTutorial() {
        var dbT = dbReadOnlyTutorial;
        if (dbT) {
            dbT.active = false;
            if (dbT.timerId != null) {
                clearInterval(dbT.timerId);
                dbT.timerId = null;
            }
            if (dbT.autoStepTimerId != null) {
                clearTimeout(dbT.autoStepTimerId);
                dbT.autoStepTimerId = null;
            }
            dbT.pendingAdvanceStepIndex = -1;
            dbT.pendingAdvanceSinceMs = 0;
        }
        var puT = pendingUnknownTutorial;
        if (puT) {
            puT.active = false;
            if (puT.timerId != null) {
                clearInterval(puT.timerId);
                puT.timerId = null;
            }
            if (puT.autoStepTimerId != null) {
                clearTimeout(puT.autoStepTimerId);
                puT.autoStepTimerId = null;
            }
            puT.pendingAdvanceStepIndex = -1;
            puT.pendingAdvanceSinceMs = 0;
        }
        setDbTutorialVisibility(false);
        hideDbTutorialCompleteOverlay();
    }

    function startDbReadOnlyTutorialIfNeeded() {
        if (
            activeTutorialMode !== EBOB_TUTORIAL_MODE_DB_READ_ONLY &&
            activeTutorialMode !== EBOB_TUTORIAL_MODE_PENDING_UNKNOWN
        ) {
            stopDbReadOnlyTutorial();
            return;
        }
        bindDbReadOnlyTutorialGuards();
        var t = guidedGetActiveTutorialState();
        if (!t || t.completed) return;
        t.active = true;
        if (t.enteredStepIndex > t.stepIndex) t.enteredStepIndex = -1;
        hideDbTutorialCompleteOverlay();
        renderDbReadOnlyTutorial();
        if (t.timerId == null) {
            t.timerId = setInterval(tickDbReadOnlyTutorial, 250);
        }
        var dom = getDbTutorialDom();
        if (dom.next && !dom.next.__ebobTutorialBound) {
            dom.next.__ebobTutorialBound = true;
            dom.next.addEventListener("click", function () {
                var tt = guidedGetActiveTutorialState();
                if (!tt) return;
                var step = tt.steps[tt.stepIndex];
                if (!step || !step.requiresNext) return;
                if (typeof step.onNext === "function") {
                    var shouldAdvance = step.onNext();
                    if (shouldAdvance === false) return;
                }
                advanceDbReadOnlyTutorial();
            });
        }
        if (dom.completeDone && !dom.completeDone.__ebobTutorialBound) {
            dom.completeDone.__ebobTutorialBound = true;
            dom.completeDone.addEventListener("click", function () {
                returnToTutorialScenarioMenu();
            });
        }
    }

    function getCmdIpv4HighlightRect() {
        var out = document.getElementById("cmdOutput");
        if (!out || !isVisibleForTutorial(out)) return null;
        var text = String(out.textContent || "");
        if (!text) return null;
        var lines = text.split("\n");
        var idx = -1;
        for (var i = 0; i < lines.length; i++) {
            if (lines[i].indexOf("IPv4 Address") >= 0) {
                idx = i;
                break;
            }
        }
        if (idx < 0) return null;
        var rect = out.getBoundingClientRect();
        var cs = window.getComputedStyle(out);
        var lh = parseFloat(cs.lineHeight);
        if (!isFinite(lh) || lh <= 0) {
            var fs = parseFloat(cs.fontSize);
            lh = isFinite(fs) && fs > 0 ? fs * 1.35 : 20;
        }
        var top = rect.top + idx * lh - 2;
        return {
            left: rect.left + 6,
            top: top,
            width: Math.max(180, rect.width - 12),
            height: lh + 4
        };
    }

    /* —— init —— */
    (function bootstrapBinventory() {
        if (wireScenarioGateIfNeeded()) return;
        loadState();
        if (
            activeTutorialMode === EBOB_TUTORIAL_MODE_DB_READ_ONLY ||
            activeTutorialMode === EBOB_TUTORIAL_MODE_PENDING_UNKNOWN
        ) {
            applyAutoLoginSessionIfEnabled();
            refreshUI();
            if (pendingImportRestartToast) {
                pendingImportRestartToast = false;
                toast("Binventory Workstation restarted after system import.");
            }
            startDbReadOnlyTutorialIfNeeded();
            return;
        }
        runBinventoryStartupSequence(function () {
            applyAutoLoginSessionIfEnabled();
            refreshUI({ staggerVesselReveal: true });
            if (pendingImportRestartToast) {
                pendingImportRestartToast = false;
                toast("Binventory Workstation restarted after system import.");
            }
            startDbReadOnlyTutorialIfNeeded();
        });
    })();
    initVesselAlarmStrobeTimer();
    wireAssignContactsDialog();
    wireSimSystemExportImport();

    (function wireGuideBackToMenu() {
        var a = document.getElementById("guideBackToMenu");
        if (!a || a.__ebobGuideBackBound) return;
        a.__ebobGuideBackBound = true;
        a.addEventListener("click", function (e) {
            e.preventDefault();
            returnToTutorialScenarioMenu();
        });
    })();

    (function wireDesktopRelaunch() {
        var icon = document.getElementById("desktopEbobIcon");
        if (icon) {
            icon.addEventListener("dblclick", function () {
                launchEbobFromDesktop();
            });
        }
        var tbEbob = document.getElementById("taskbarBtnEbob");
        if (tbEbob) {
            tbEbob.addEventListener("click", function () {
                var desk = document.getElementById("simDesktop");
                var pw = document.getElementById("pageWrap");
                var shell = document.getElementById("appShell");
                if (pw && pw.classList.contains("page-wrap--desktop-mode")) {
                    if (desk && !desk.hidden) launchEbobFromDesktop();
                    return;
                }
                if (shell && shell.classList.contains("app-shell--minimized")) {
                    shell.classList.remove("app-shell--minimized");
                    tbEbob.classList.remove("win-taskbar-ebob--minimized");
                    return;
                }
                if (shell && shell.scrollIntoView) shell.scrollIntoView({ behavior: "smooth", block: "nearest" });
            });
        }
    })();

    (function wireTaskbarJumpMenu() {
        var menu = document.getElementById("taskbarJumpList");
        var closeBtn = document.getElementById("taskbarJumpClose");
        var target = null;

        function hideJump() {
            if (!menu) return;
            menu.hidden = true;
            menu.setAttribute("aria-hidden", "true");
            target = null;
        }

        function showJump(clientX, clientY, which) {
            if (!menu || !closeBtn) return;
            target = which;
            menu.hidden = false;
            menu.setAttribute("aria-hidden", "false");
            menu.style.left = Math.min(clientX, window.innerWidth - 210) + "px";
            menu.style.top = Math.min(clientY, window.innerHeight - 48) + "px";
            setTimeout(function () {
                closeBtn.focus();
            }, 0);
        }

        if (closeBtn) {
            closeBtn.addEventListener("click", function () {
                if (target === "ebob") exitToSimDesktop();
                else if (target === "services") {
                    if (typeof window.__ebobCloseServicesMsc === "function") window.__ebobCloseServicesMsc();
                }
                hideJump();
            });
        }

        document.addEventListener(
            "click",
            function (e) {
                if (!menu || menu.hidden) return;
                if (e.target.closest && e.target.closest("#taskbarJumpList")) return;
                hideJump();
            },
            true
        );

        document.addEventListener(
            "contextmenu",
            function (e) {
                var eb = e.target.closest && e.target.closest("#taskbarBtnEbob");
                var sv = e.target.closest && e.target.closest("#taskbarBtnServices");
                if (eb) {
                    e.preventDefault();
                    showJump(e.clientX, e.clientY, "ebob");
                    return;
                }
                if (sv && !sv.hidden) {
                    e.preventDefault();
                    showJump(e.clientX, e.clientY, "services");
                }
            },
            true
        );
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
                toast("Workstation reloaded — site: " + (s && s.name ? s.name : siteId) + ".");
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
            ensureVesselFields(v, 0);
            var vw = formatVolumeWeightPair(v, v.volumeCuFt, v.weightLb);
            return (
                "<tr><td>" +
                escapeHtml(v.name) +
                "</td><td>" +
                escapeHtml(String(v.pctFull != null ? v.pctFull : "")) +
                "%</td><td>" +
                escapeHtml(vw.volStr) +
                "</td><td>" +
                escapeHtml(vw.wtStr) +
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
        var pwInput = document.getElementById("pw");
        var rawUid = uidInput ? uidInput.value.trim() : "";
        var rawPw = pwInput ? pwInput.value : "";
        state.users.forEach(ensureUserMaintenanceUser);
        var match = null;
        for (var ui = 0; ui < state.users.length; ui++) {
            var cand = state.users[ui];
            if (String(cand.userId).toLowerCase() === rawUid.toLowerCase()) {
                match = cand;
                break;
            }
        }
        if (!match) {
            showBinventoryMessageBox({
                icon: "warn",
                message: "Invalid User ID or password.",
                buttons: "ok"
            });
            return;
        }
        if (match.authenticationMethod === 1) {
            /* LDAP — workstation validates against directory; emulator skips local password. */
        } else {
            var expected = match.password != null ? String(match.password) : "";
            if (expected !== rawPw) {
                showBinventoryMessageBox({
                    icon: "warn",
                    message: "Invalid User ID or password.",
                    buttons: "ok"
                });
                return;
            }
        }
        state.currentUser = match.name || match.userId;
        match.lastLogon = new Date().toLocaleString();
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
                showLoginBackdrop();
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
        else if (a === "minimize") {
            var shell = document.getElementById("appShell");
            var tbEbob = document.getElementById("taskbarBtnEbob");
            if (shell) shell.classList.add("app-shell--minimized");
            if (tbEbob) tbEbob.classList.add("win-taskbar-ebob--minimized");
        }
    });

    document.getElementById("appShell").addEventListener("click", function (e) {
        var siloEl = e.target.closest(".silo");
        if (siloEl) {
            var card0 = siloEl.closest(".vessel");
            var id0 = card0 ? card0.dataset.vesselId : null;
            var v0 = id0 ? findVessel(id0) : null;
            if (v0) {
                e.preventDefault();
                e.stopPropagation();
                openVesselDetails(v0);
            }
            return;
        }
        var a = e.target.closest("[data-vessel-action]");
        if (!a) return;
        var act = a.getAttribute("data-vessel-action");
        var card = a.closest(".vessel");
        var id = card ? card.dataset.vesselId : null;
        var v = id ? findVessel(id) : null;
        if (act === "measure" && v) measureVessel(v);
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
            var vsActive =
                vmSetupEditingId &&
                (appModalShell.classList.contains("modal-vessel-setup") ||
                    (document.getElementById("backdropAppStack") &&
                        document.getElementById("backdropAppStack").classList.contains("show") &&
                        document.getElementById("appModalShellStack") &&
                        document.getElementById("appModalShellStack").classList.contains("modal-vessel-setup")));
            if (bdApp.classList.contains("show") && vsActive) {
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

    /* Modals close only via caption ✕ / Cancel / Save — not by clicking the dimmed backdrop (WinForms-style). */

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
            var lastScheduleHm = "";
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
                var hm = pad2Sched(now.getHours()) + ":" + pad2Sched(now.getMinutes());
                if (hm !== lastScheduleHm) {
                    lastScheduleHm = hm;
                    processEmulatedScheduleTick(now);
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
        var winSearchPanelCmd = document.getElementById("winSearchPanelCmd");
        var winSearchPanelServices = document.getElementById("winSearchPanelServices");
        var winSearchPanelDevMgr = document.getElementById("winSearchPanelDevMgr");
        var winStartHomeContent = document.getElementById("winStartHomeContent");
        /** Full ipconfig-style output (no global Primary Dns line); Wi-Fi uses home.local for connection-specific suffix. */
        var CMD_IPCONFIG_OUTPUT = [
            "IP configuration",
            "",
            "",
            "Ethernet adapter Ethernet 3:",
            "",
            "   Media State . . . . . . . . . . . : Media disconnected",
            "   Connection-specific DNS Suffix  . : ",
            "",
            "Ethernet adapter Ethernet:",
            "",
            "   Connection-specific DNS Suffix  . : YOURCOMPANY.LOCAL",
            "   Link-local IPv6 Address . . . . . : fe80::5970:1669:1b50:c4d%6",
            "   IPv4 Address. . . . . . . . . . . : " + SIM_WORKSTATION_IPV4,
            "   Subnet Mask . . . . . . . . . . . : 255.255.254.0",
            "   Default Gateway . . . . . . . . . : 10.101.70.1",
            "",
            "Wireless LAN adapter Wi-Fi:",
            "",
            "   Media State . . . . . . . . . . . : Media disconnected",
            "   Connection-specific DNS Suffix  . : home.local",
            "",
            "Wireless LAN adapter Local Area Connection* 9:",
            "",
            "   Media State . . . . . . . . . . . : Media disconnected",
            "   Connection-specific DNS Suffix  . : ",
            "",
            "Wireless LAN adapter Local Area Connection* 10:",
            "",
            "   Media State . . . . . . . . . . . : Media disconnected",
            "   Connection-specific DNS Suffix  . : ",
            "",
            "Ethernet adapter Bluetooth Network Connection:",
            "",
            "   Media State . . . . . . . . . . . : Media disconnected",
            "   Connection-specific DNS Suffix  . : ",
            ""
        ].join("\n");

        function openSimCommandPrompt() {
            var body =
                '<div class="cmd-shell" id="cmdShell">' +
                '<div class="cmd-scroll" id="cmdScroll" role="log" aria-live="polite">' +
                '<div class="cmd-banner" id="cmdBanner"></div>' +
                '<pre class="cmd-output" id="cmdOutput"></pre>' +
                '<div class="cmd-line" id="cmdLine">' +
                '<label class="cmd-input-row" for="cmdInput">' +
                '<span class="cmd-prompt">H:\\&gt;</span>' +
                '<input type="text" id="cmdInput" class="cmd-input" autocomplete="off" spellcheck="false" aria-label="Command input">' +
                "</label>" +
                "</div>" +
                "</div>" +
                "</div>";
            openAppModal("Terminal", body, "", "modal-command-prompt modal-win-toolwindow");

            var cmdShell = document.getElementById("cmdShell");
            var cmdScroll = document.getElementById("cmdScroll");
            var cmdBanner = document.getElementById("cmdBanner");
            var cmdOutput = document.getElementById("cmdOutput");
            var cmdInput = document.getElementById("cmdInput");
            if (!cmdOutput || !cmdInput) return;

            function selectionInsideCmdShell() {
                try {
                    var s = window.getSelection();
                    if (!s || !s.rangeCount || s.isCollapsed) return false;
                    return cmdShell.contains(s.getRangeAt(0).commonAncestorContainer);
                } catch (err) {
                    return false;
                }
            }

            function canCopyFromCmdUi() {
                if (
                    document.activeElement === cmdInput &&
                    typeof cmdInput.selectionStart === "number" &&
                    cmdInput.selectionStart !== cmdInput.selectionEnd
                ) {
                    return true;
                }
                return selectionInsideCmdShell();
            }

            function copyTextWithFallback(text) {
                if (!text) return Promise.resolve(false);
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    return navigator.clipboard
                        .writeText(text)
                        .then(function () {
                            return true;
                        })
                        .catch(function () {
                            return Promise.resolve(legacyCopyText(text));
                        });
                }
                return Promise.resolve(legacyCopyText(text));
            }

            function legacyCopyText(text) {
                var activeEl = document.activeElement;
                var selection = window.getSelection ? window.getSelection() : null;
                var savedRanges = [];
                if (selection && selection.rangeCount) {
                    for (var i = 0; i < selection.rangeCount; i++) {
                        savedRanges.push(selection.getRangeAt(i).cloneRange());
                    }
                }
                var ta = document.createElement("textarea");
                ta.setAttribute("readonly", "");
                ta.value = text;
                ta.style.position = "fixed";
                ta.style.left = "-9999px";
                document.body.appendChild(ta);
                ta.select();
                try {
                    return document.execCommand("copy");
                } catch (err) {
                    return false;
                } finally {
                    document.body.removeChild(ta);
                    if (selection) {
                        selection.removeAllRanges();
                        for (var r = 0; r < savedRanges.length; r++) {
                            selection.addRange(savedRanges[r]);
                        }
                    }
                    if (activeEl && activeEl.focus) activeEl.focus();
                }
            }

            /** Text selected in the transcript or in the prompt input (getSelection() misses many input selections). */
            function getSelectedTextForCmd() {
                if (
                    document.activeElement === cmdInput &&
                    typeof cmdInput.selectionStart === "number" &&
                    cmdInput.selectionStart !== cmdInput.selectionEnd
                ) {
                    return String(cmdInput.value || "").slice(cmdInput.selectionStart, cmdInput.selectionEnd);
                }
                var s = window.getSelection && window.getSelection();
                return s ? s.toString() : "";
            }

            function pasteIntoCmdInput(text) {
                if (text == null || text === "") return;
                cmdInput.focus();
                var v = String(cmdInput.value || "");
                var start =
                    typeof cmdInput.selectionStart === "number" ? cmdInput.selectionStart : v.length;
                var end = typeof cmdInput.selectionEnd === "number" ? cmdInput.selectionEnd : start;
                cmdInput.value = v.slice(0, start) + text + v.slice(end);
                var pos = start + String(text).length;
                if (typeof cmdInput.setSelectionRange === "function") {
                    cmdInput.setSelectionRange(pos, pos);
                }
            }

            function clearCmdTextSelection() {
                var s = window.getSelection && window.getSelection();
                if (s) s.removeAllRanges();
                if (document.activeElement === cmdInput && typeof cmdInput.selectionStart === "number") {
                    var p = cmdInput.selectionStart;
                    cmdInput.setSelectionRange(p, p);
                }
            }

            if (cmdShell) {
                cmdShell.addEventListener("click", function (e) {
                    if (cmdInput.contains(e.target)) return;
                    if (e.target.closest && (e.target.closest("#cmdOutput") || e.target.closest("#cmdBanner"))) {
                        return;
                    }
                    cmdInput.focus();
                });

                cmdShell.addEventListener("contextmenu", function (e) {
                    e.preventDefault();
                    var t = getSelectedTextForCmd();
                    if (t && t.trim()) {
                        copyTextWithFallback(t).then(function (ok) {
                            if (ok) {
                                toast("Copied to clipboard.");
                                clearCmdTextSelection();
                                cmdInput.focus();
                            }
                        });
                        return;
                    }
                    if (navigator.clipboard && navigator.clipboard.readText) {
                        navigator.clipboard.readText().then(pasteIntoCmdInput).catch(function () {});
                    }
                });

                cmdShell.addEventListener(
                    "keydown",
                    function (e) {
                        if (!e.ctrlKey || String(e.key).toLowerCase() !== "c") return;
                        var t = getSelectedTextForCmd();
                        if (!t || !t.trim()) return;
                        if (!canCopyFromCmdUi()) return;
                        e.preventDefault();
                        copyTextWithFallback(t).then(function (ok) {
                            if (ok) toast("Copied to clipboard.");
                        });
                    },
                    true
                );
            }

            if (cmdBanner) {
                cmdBanner.textContent = [
                    "Terminal Shell [Version 5.4.9 build 8117]",
                    "Copyright (c) simulation"
                ].join("\n");
            }

            function scrollCmdToBottom() {
                if (cmdScroll) cmdScroll.scrollTop = cmdScroll.scrollHeight;
            }

            function appendLine(line) {
                if (cmdOutput.textContent.length > 0) cmdOutput.textContent += "\n";
                cmdOutput.textContent += line;
                scrollCmdToBottom();
            }

            scrollCmdToBottom();

            cmdInput.addEventListener("keydown", function (evt) {
                if (evt.key !== "Enter") return;
                evt.preventDefault();
                var raw = String(cmdInput.value || "");
                var cmd = raw.trim().toLowerCase();
                appendLine("H:\\>" + raw);
                if (cmd === "cls") {
                    cmdOutput.textContent = "";
                    if (cmdBanner) cmdBanner.textContent = "";
                    cmdInput.value = "";
                    scrollCmdToBottom();
                    return;
                }
                appendLine("");
                if (cmd === "ipconfig") {
                    appendLine(CMD_IPCONFIG_OUTPUT);
                } else if (cmd.length > 0) {
                    appendLine("'" + raw + "' is not recognized as an internal or external command,");
                    appendLine("operable program or batch file.");
                }
                appendLine("");
                cmdInput.value = "";
            });

            setTimeout(function () {
                cmdInput.focus();
            }, 0);
        }

        function syncStartSearchView() {
            if (!winSearchFlyout || !winStartHomeContent || !startMenu) return;
            var q = String(winStartSearch && winStartSearch.value ? winStartSearch.value : "").trim().toLowerCase();
            var isCmd =
                q.indexOf("cmd") === 0 ||
                q === "command prompt" ||
                q.indexOf("terminal") === 0 ||
                q === "terminal";
            var isServices = q === "services" || q.indexOf("services") === 0;
            var isDevMgr =
                q === "device" ||
                q === "devices" ||
                q.indexOf("device ") === 0 ||
                q.indexOf("devices ") === 0 ||
                q.indexOf("device manager") === 0 ||
                q === "devmgmt" ||
                q.indexOf("devmgmt") === 0;
            if (isCmd || isServices || isDevMgr) {
                winSearchFlyout.hidden = false;
                winStartHomeContent.hidden = true;
                startMenu.classList.add("win-start-menu--search");
                if (winSearchPanelCmd) winSearchPanelCmd.hidden = !isCmd;
                if (winSearchPanelServices) winSearchPanelServices.hidden = !isServices;
                if (winSearchPanelDevMgr) winSearchPanelDevMgr.hidden = !isDevMgr;
            } else {
                winSearchFlyout.hidden = true;
                winStartHomeContent.hidden = false;
                startMenu.classList.remove("win-start-menu--search");
                if (winSearchPanelCmd) winSearchPanelCmd.hidden = true;
                if (winSearchPanelServices) winSearchPanelServices.hidden = true;
                if (winSearchPanelDevMgr) winSearchPanelDevMgr.hidden = true;
            }
        }

        function closeStartMenu() {
            if (!backdrop || !startMenu || !startBtn) return;
            if (winStartSearch && document.activeElement === winStartSearch) {
                winStartSearch.blur();
            }
            backdrop.hidden = true;
            backdrop.setAttribute("aria-hidden", "true");
            startMenu.hidden = true;
            startBtn.setAttribute("aria-expanded", "false");
            if (winStartSearch) winStartSearch.value = "";
            syncStartSearchView();
            /* Do not focus the Start button — that drew an obvious focus ring after Enter-to-launch (unlike real Windows). */
            if (document.activeElement === startBtn) {
                startBtn.blur();
            }
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
            winStartSearch.addEventListener("keydown", function (e) {
                if (e.key !== "Enter") return;
                var q = String(winStartSearch.value || "").trim().toLowerCase();
                if (
                    q.indexOf("cmd") === 0 ||
                    q === "command prompt" ||
                    q.indexOf("terminal") === 0 ||
                    q === "terminal"
                ) {
                    e.preventDefault();
                    closeStartMenu();
                    openSimCommandPrompt();
                    return;
                }
                if (q === "services" || q.indexOf("services") === 0) {
                    e.preventDefault();
                    closeStartMenu();
                    if (typeof window.__ebobShowServicesMsc === "function") {
                        window.__ebobShowServicesMsc();
                    }
                    return;
                }
                var isDevMgrEnter =
                    q === "device" ||
                    q === "devices" ||
                    q.indexOf("device ") === 0 ||
                    q.indexOf("devices ") === 0 ||
                    q.indexOf("device manager") === 0 ||
                    q === "devmgmt" ||
                    q.indexOf("devmgmt") === 0;
                if (isDevMgrEnter) {
                    e.preventDefault();
                    closeStartMenu();
                    var dmBdS = document.getElementById("backdropDeviceMgr");
                    if (
                        dmBdS &&
                        dmBdS.classList.contains("backdrop-device-mgr--minimized") &&
                        typeof window.__ebobRestoreDeviceMgr === "function"
                    ) {
                        window.__ebobRestoreDeviceMgr();
                        return;
                    }
                    if (typeof window.__ebobShowDeviceMgr === "function") {
                        window.__ebobShowDeviceMgr();
                    }
                }
            });
        }

        if (startMenu) {
            startMenu.addEventListener("click", function (e) {
                var t = e.target;
                var dmgBtn = t.closest && t.closest("[data-sim-dmg]");
                if (dmgBtn) {
                    var dact = dmgBtn.getAttribute("data-sim-dmg");
                    if (dact === "open") {
                        closeStartMenu();
                        var dmBdOpen = document.getElementById("backdropDeviceMgr");
                        if (
                            dmBdOpen &&
                            dmBdOpen.classList.contains("backdrop-device-mgr--minimized") &&
                            typeof window.__ebobRestoreDeviceMgr === "function"
                        ) {
                            window.__ebobRestoreDeviceMgr();
                            return;
                        }
                        if (typeof window.__ebobShowDeviceMgr === "function") {
                            window.__ebobShowDeviceMgr();
                        }
                        return;
                    }
                    if (dact === "admin") {
                        toast("Device Manager (elevated) would open here.");
                        return;
                    }
                    if (dact === "loc") {
                        toast("Control Panel location opened.");
                        return;
                    }
                    if (dact === "pstart") {
                        toast("Device Manager pinned to Start.");
                        return;
                    }
                    if (dact === "ptask") {
                        toast("Device Manager pinned to taskbar.");
                        return;
                    }
                    return;
                }
                var svcBtn = t.closest && t.closest("[data-sim-svc]");
                if (svcBtn) {
                    var sact = svcBtn.getAttribute("data-sim-svc");
                    if (sact === "open") {
                        closeStartMenu();
                        if (typeof window.__ebobShowServicesMsc === "function") {
                            window.__ebobShowServicesMsc();
                        }
                        return;
                    }
                    if (sact === "admin") {
                        closeStartMenu();
                        var svcBdMin = document.getElementById("backdropServicesMsc");
                        if (
                            svcBdMin &&
                            svcBdMin.classList.contains("backdrop-services--minimized") &&
                            typeof window.__ebobRestoreServicesMsc === "function"
                        ) {
                            window.__ebobRestoreServicesMsc();
                            return;
                        }
                        if (typeof window.__ebobShowServicesUac === "function") {
                            window.__ebobShowServicesUac();
                        }
                        return;
                    }
                    return;
                }
                var cmdBtn = t.closest && t.closest("[data-sim-cmd]");
                if (cmdBtn) {
                    var act = cmdBtn.getAttribute("data-sim-cmd");
                    if (act === "open" || act === "admin") {
                        closeStartMenu();
                        var bdAppCmd = document.getElementById("backdropApp");
                        if (
                            bdAppCmd &&
                            bdAppCmd.classList.contains("backdrop-app--terminal-minimized") &&
                            typeof window.__ebobRestoreTerminalApp === "function"
                        ) {
                            window.__ebobRestoreTerminalApp();
                            return;
                        }
                        openSimCommandPrompt();
                        return;
                    }
                    if (act === "loc") {
                        toast("Terminal location opened.");
                        return;
                    }
                    if (act === "pstart") {
                        toast("Terminal pinned to Start.");
                        return;
                    }
                    if (act === "ptask") {
                        toast("Terminal pinned to taskbar.");
                        return;
                    }
                    return;
                }
                if (t.closest && t.closest("#winSearchHitCmd")) {
                    closeStartMenu();
                    var bdAppHitCmd = document.getElementById("backdropApp");
                    if (
                        bdAppHitCmd &&
                        bdAppHitCmd.classList.contains("backdrop-app--terminal-minimized") &&
                        typeof window.__ebobRestoreTerminalApp === "function"
                    ) {
                        window.__ebobRestoreTerminalApp();
                        return;
                    }
                    openSimCommandPrompt();
                    return;
                }
                if (t.closest && t.closest("#winSearchHitServices")) {
                    closeStartMenu();
                    if (typeof window.__ebobShowServicesMsc === "function") {
                        window.__ebobShowServicesMsc();
                    }
                    return;
                }
                if (t.closest && t.closest("#winSearchHitDevMgr")) {
                    closeStartMenu();
                    var dmBdHit = document.getElementById("backdropDeviceMgr");
                    if (
                        dmBdHit &&
                        dmBdHit.classList.contains("backdrop-device-mgr--minimized") &&
                        typeof window.__ebobRestoreDeviceMgr === "function"
                    ) {
                        window.__ebobRestoreDeviceMgr();
                        return;
                    }
                    if (typeof window.__ebobShowDeviceMgr === "function") {
                        window.__ebobShowDeviceMgr();
                    }
                    return;
                }
                var tile = t.closest && t.closest(".win-start-tile[data-sim-app]");
                if (tile) {
                    closeStartMenu();
                    return;
                }
                var rec = t.closest && t.closest(".win-start-rec button[data-sim-rec]");
                if (rec) {
                    var k = rec.getAttribute("data-sim-rec");
                    if (k === "ebob") {
                        window.scrollTo({ top: 0, behavior: "smooth" });
                    }
                    closeStartMenu();
                    return;
                }
                var hdrLink = t.closest && t.closest("[data-sim-start]");
                if (hdrLink) {
                    closeStartMenu();
                    return;
                }
            });
        }

    })();

    (function wireDeviceManagerWin() {
        var backdropDm = document.getElementById("backdropDeviceMgr");
        var shellDm = document.getElementById("deviceMgrShell");
        var dmClose = document.getElementById("dmMscClose");
        var dmMin = document.getElementById("dmMscMin");
        var dmMax = document.getElementById("dmMscMax");
        var taskbarBtnDeviceMgr = document.getElementById("taskbarBtnDeviceMgr");

        function syncDeviceMgrTaskbar(visible) {
            if (!taskbarBtnDeviceMgr) return;
            if (visible) {
                taskbarBtnDeviceMgr.hidden = false;
                taskbarBtnDeviceMgr.classList.add("win-taskbar-device-mgr--active");
            } else {
                taskbarBtnDeviceMgr.hidden = true;
                taskbarBtnDeviceMgr.classList.remove("win-taskbar-device-mgr--active");
            }
        }

        function minimizeDeviceMgr() {
            if (!backdropDm) return;
            backdropDm.classList.remove("show");
            backdropDm.classList.add("backdrop-device-mgr--minimized");
            backdropDm.setAttribute("aria-hidden", "true");
            syncDeviceMgrTaskbar(true);
        }

        function restoreDeviceMgr() {
            if (!backdropDm) return;
            backdropDm.classList.remove("backdrop-device-mgr--minimized");
            backdropDm.classList.add("show");
            backdropDm.setAttribute("aria-hidden", "false");
            syncDeviceMgrTaskbar(true);
        }

        /** KDE Breeze icons (LGPL) — see assets/device-mgr/ */
        var DM_ICON_BASE = "assets/device-mgr/";

        var DEVICE_MGR_ROWS = [
            { label: "Audio inputs and outputs", icon: "audio-headphones.svg" },
            { label: "Audio Processing Objects (APOs)", icon: "audio-card.svg" },
            { label: "Batteries", icon: "phone-battery.svg" },
            { label: "Bluetooth", icon: "network-bluetooth.svg" },
            { label: "Cameras", icon: "camera-web.svg" },
            { label: "Computer", icon: "computer.svg" },
            { label: "Disk drives", icon: "drive-harddisk.svg" },
            { label: "Display adapters", icon: "video-television.svg" },
            { label: "Firmware", icon: "preferences-devices-cpu.svg" },
            { label: "Human Interface Devices", icon: "input-tablet.svg" },
            { label: "Keyboards", icon: "input-keyboard.svg" },
            { label: "Mice and other pointing devices", icon: "input-mouse.svg" },
            { label: "Monitors", icon: "monitor.svg" },
            { label: "Network adapters", icon: "network-wired.svg" },
            { label: "Ports (COM & LPT)", icon: "network-modem.svg", ports: true },
            { label: "Print queues", icon: "document-print.svg" },
            { label: "Printers", icon: "printer.svg" },
            { label: "Processors", icon: "preferences-devices-cpu.svg" },
            { label: "Security devices", icon: "preferences-security.svg" },
            { label: "Software components", icon: "applications-other.svg" },
            { label: "Software devices", icon: "preferences-plugin.svg" },
            { label: "Sound, video and game controllers", icon: "applications-multimedia.svg" },
            { label: "Storage controllers", icon: "preferences-system-disks.svg" },
            { label: "System devices", icon: "applications-system.svg" },
            { label: "Universal Serial Bus controllers", icon: "drive-removable-media-usb.svg" },
            { label: "Universal Serial Bus devices", icon: "drive-removable-media-usb-pendrive.svg" },
            { label: "USB Connector Managers", icon: "network-rj45-female.svg" }
        ];

        var DEVICE_MGR_PORT_CHILDREN = [
            "RS-485 Isolated Port (COM2)",
            "Standard Serial over Bluetooth link (COM3)",
            "Standard Serial over Bluetooth link (COM4)",
            "Standard Serial over Bluetooth link (COM16)",
            "Standard Serial over Bluetooth link (COM17)"
        ];

        function dmCategoryIconHtml(iconFile) {
            return (
                '<span class="dm-cat-ico" aria-hidden="true">' +
                '<img class="dm-cat-ico-img" src="' +
                DM_ICON_BASE +
                iconFile +
                '" width="16" height="16" alt="" loading="lazy" decoding="async">' +
                "</span>"
            );
        }

        function dmPortLeafIconHtml() {
            return (
                '<span class="dm-port-ico" aria-hidden="true">' +
                '<img class="dm-port-ico-img" src="' +
                DM_ICON_BASE +
                'network-modem.svg" width="14" height="14" alt="" loading="lazy" decoding="async">' +
                "</span>"
            );
        }

        function ensureDeviceMgrTreeBuilt() {
            var host = document.getElementById("deviceMgrCatList");
            if (!host || host.getAttribute("data-built") === "1") return;

            var parts = [];
            for (var i = 0; i < DEVICE_MGR_ROWS.length; i++) {
                var row = DEVICE_MGR_ROWS[i];
                if (row.ports) {
                    parts.push(
                        '<div class="dm-cat dm-cat--ports" role="presentation">' +
                            '<button type="button" class="dm-ports-chev" id="devmgrPortsToggle" aria-expanded="false" aria-controls="devmgrPortsChildren" title="Expand or collapse">›</button>' +
                            dmCategoryIconHtml(row.icon) +
                            '<span class="dm-cat-lbl">Ports (COM &amp; LPT)</span>' +
                            "</div>" +
                            '<div class="dm-ports-kids" id="devmgrPortsChildren" role="group" aria-label="Serial ports" hidden>'
                    );
                    for (var p = 0; p < DEVICE_MGR_PORT_CHILDREN.length; p++) {
                        parts.push(
                            '<div class="dm-port-leaf" role="treeitem">' +
                                dmPortLeafIconHtml() +
                                '<span class="dm-port-lbl">' +
                                escapeHtml(DEVICE_MGR_PORT_CHILDREN[p]) +
                                "</span></div>"
                        );
                    }
                    parts.push("</div>");
                } else {
                    parts.push(
                        '<div class="dm-cat dm-cat--nogrow" role="treeitem" aria-expanded="false">' +
                            '<span class="dm-chev-fake" aria-hidden="true">›</span>' +
                            dmCategoryIconHtml(row.icon) +
                            '<span class="dm-cat-lbl">' +
                            escapeHtml(row.label) +
                            "</span></div>"
                    );
                }
            }
            host.innerHTML = parts.join("");
            host.setAttribute("data-built", "1");

            var btn = document.getElementById("devmgrPortsToggle");
            var kids = document.getElementById("devmgrPortsChildren");
            if (btn && kids) {
                btn.addEventListener("click", function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    kids.hidden = !kids.hidden;
                    var isOpen = !kids.hidden;
                    btn.setAttribute("aria-expanded", isOpen ? "true" : "false");
                    btn.classList.toggle("dm-ports-chev--open", isOpen);
                });
            }
        }

        function showDeviceManager() {
            if (!backdropDm) return;
            ensureDeviceMgrTreeBuilt();
            var kids = document.getElementById("devmgrPortsChildren");
            var btn = document.getElementById("devmgrPortsToggle");
            if (kids) kids.hidden = true;
            if (btn) {
                btn.setAttribute("aria-expanded", "false");
                btn.classList.remove("dm-ports-chev--open");
            }
            backdropDm.classList.remove("backdrop-device-mgr--minimized");
            backdropDm.classList.add("show");
            backdropDm.setAttribute("aria-hidden", "false");
            syncDeviceMgrTaskbar(true);
        }

        function hideDeviceManager() {
            if (!backdropDm) return;
            backdropDm.classList.remove("show", "backdrop-device-mgr--minimized");
            backdropDm.setAttribute("aria-hidden", "true");
            if (shellDm) shellDm.classList.remove("modal-device-mgr--max");
            syncDeviceMgrTaskbar(false);
        }

        function onDmKeydown(e) {
            if (!backdropDm || !backdropDm.classList.contains("show")) return;
            var tag = e.target && e.target.tagName;
            if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
            if (e.key === "Escape") {
                e.preventDefault();
                e.stopPropagation();
                minimizeDeviceMgr();
            }
        }

        document.addEventListener("keydown", onDmKeydown, true);

        if (dmClose) dmClose.addEventListener("click", hideDeviceManager);
        if (dmMin) dmMin.addEventListener("click", minimizeDeviceMgr);
        if (dmMax && shellDm) {
            dmMax.addEventListener("click", function () {
                shellDm.classList.toggle("modal-device-mgr--max");
            });
        }

        if (taskbarBtnDeviceMgr) {
            taskbarBtnDeviceMgr.addEventListener("click", function () {
                if (!backdropDm) return;
                if (backdropDm.classList.contains("show")) return;
                restoreDeviceMgr();
            });
        }

        window.__ebobShowDeviceMgr = showDeviceManager;
        window.__ebobCloseDeviceMgr = hideDeviceManager;
        window.__ebobRestoreDeviceMgr = restoreDeviceMgr;
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

        function getSvcMscTableEl() {
            return backdropSvc ? backdropSvc.querySelector("#svcMscTable") : document.getElementById("svcMscTable");
        }

        /** Always target the tbody inside the Services backdrop (avoids wrong table / empty list). */
        function getSvcMscTbody() {
            if (backdropSvc) {
                var t = backdropSvc.querySelector("tbody#svcMscTbody");
                if (t) return t;
                var tbl = backdropSvc.querySelector("table#svcMscTable");
                if (tbl) {
                    var tb = tbl.querySelector("tbody");
                    if (tb) return tb;
                }
            }
            var fallback = document.getElementById("svcMscTbody");
            return fallback || null;
        }

        var uacYes = document.getElementById("uacYes");
        var uacNo = document.getElementById("uacNo");
        var uacClose = document.getElementById("uacClose");
        var uacMoreDetails = document.getElementById("uacMoreDetails");
        var uacDetails = document.getElementById("uacDetails");
        var svcMscClose = document.getElementById("svcMscClose");
        var svcMscMin = document.getElementById("svcMscMin");
        var svcMscMax = document.getElementById("svcMscMax");
        var servicesMscShell = document.getElementById("servicesMscShell");
        var taskbarBtnServices = document.getElementById("taskbarBtnServices");
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
            if (uacDetails && uacMoreDetails) {
                uacDetails.hidden = true;
                uacMoreDetails.setAttribute("aria-expanded", "false");
                uacMoreDetails.textContent = "Show more details";
            }
        }

        function syncMscEngineToWorkstation() {
            var eng = getService("ebob-engine");
            var sch = getService("ebob-scheduler");
            if (eng && sch && !eng.running && sch.running) {
                sch.running = false;
                applyStatusToRow("ebob-scheduler");
            }
            if (sch) state.ebobSchedulerRunning = sch.running;
            if (!eng) return;
            setEbobServicesRunningFromSim(eng.running);
        }

        function minimizeServicesMsc() {
            if (!backdropSvc) return;
            backdropSvc.classList.remove("show");
            backdropSvc.classList.add("backdrop-services--minimized");
            backdropSvc.setAttribute("aria-hidden", "true");
            hideCtxMenu();
            /* Like Windows: minimized window stays on the taskbar until closed. */
            if (taskbarBtnServices) {
                taskbarBtnServices.hidden = false;
                taskbarBtnServices.classList.add("win-taskbar-services--active");
            }
        }

        function restoreServicesMsc() {
            if (!backdropSvc) return;
            backdropSvc.classList.remove("backdrop-services--minimized");
            backdropSvc.classList.add("show");
            backdropSvc.setAttribute("aria-hidden", "false");
            if (taskbarBtnServices) {
                taskbarBtnServices.hidden = false;
                taskbarBtnServices.classList.add("win-taskbar-services--active");
            }
            setTimeout(function () {
                var wrap = document.querySelector(".svc-msc-table-wrap");
                if (wrap) wrap.focus();
            }, 0);
        }

        function showServicesMsc() {
            if (!backdropSvc) return;
            mscState.services = cloneSeed();
            if (!state.ebobServicesRunning) state.ebobSchedulerRunning = false;
            var engSync = getService("ebob-engine");
            var schSync = getService("ebob-scheduler");
            if (engSync) engSync.running = !!state.ebobServicesRunning;
            if (schSync) {
                schSync.running = !!state.ebobSchedulerRunning && !!state.ebobServicesRunning;
            }
            mscState.selectedId = mscState.services.length ? mscState.services[0].id : null;
            mscState.eBobCycle = 0;
            renderServicesTable();
            updateDescAndToolbar();
            backdropSvc.classList.remove("backdrop-services--minimized");
            backdropSvc.classList.add("show");
            backdropSvc.setAttribute("aria-hidden", "false");
            if (taskbarBtnServices) {
                taskbarBtnServices.hidden = false;
                taskbarBtnServices.classList.add("win-taskbar-services--active");
            }
            hideCtxMenu();
            setTimeout(function () {
                var wrap = document.querySelector(".svc-msc-table-wrap");
                if (wrap) wrap.focus();
            }, 0);
        }

        function hideServicesMsc() {
            if (!backdropSvc) return;
            backdropSvc.classList.remove("show", "backdrop-services--minimized");
            backdropSvc.setAttribute("aria-hidden", "true");
            if (servicesMscShell) servicesMscShell.classList.remove("modal-services-msc--max");
            hideCtxMenu();
            syncMscEngineToWorkstation();
            if (taskbarBtnServices) {
                taskbarBtnServices.hidden = true;
                taskbarBtnServices.classList.remove("win-taskbar-services--active");
            }
        }

        window.__ebobCloseServicesMsc = hideServicesMsc;
        window.__ebobRestoreServicesMsc = restoreServicesMsc;
        window.__ebobShowServicesMsc = showServicesMsc;

        function statusText(s) {
            return s.running ? "Running" : "";
        }

        function renderServicesTable() {
            var tb = getSvcMscTbody();
            if (!tb) return;
            tb.innerHTML = "";
            for (var i = 0; i < mscState.services.length; i++) {
                var s = mscState.services[i];
                var tr = document.createElement("tr");
                tr.className = "svc-row" + (mscState.selectedId === s.id ? " svc-row-selected" : "");
                tr.setAttribute("data-service-id", s.id);
                tr.setAttribute("role", "row");
                var tdName = document.createElement("td");
                var spanName = document.createElement("span");
                spanName.className = "svc-msc-name-text";
                spanName.textContent = s.name;
                tdName.appendChild(spanName);
                var tdDesc = document.createElement("td");
                var spanDesc = document.createElement("span");
                spanDesc.className = "svc-msc-desc-text";
                spanDesc.textContent = s.description;
                tdDesc.appendChild(spanDesc);
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
                tb.appendChild(tr);
            }
        }

        function syncRowSelectionClass() {
            var tbodyEl = getSvcMscTbody();
            var rows = tbodyEl ? tbodyEl.querySelectorAll(".svc-row") : [];
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
            var tbodyEl = getSvcMscTbody();
            if (!tbodyEl) return;
            var row = tbodyEl.querySelector('[data-service-id="' + id + '"]');
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
            var tbodyEl = getSvcMscTbody();
            var row = tbodyEl ? tbodyEl.querySelector('[data-service-id="' + id + '"]') : null;
            if (row) {
                var cells = row.querySelectorAll("td");
                if (cells.length > 2) cells[2].textContent = statusText(s);
            }
        }

        /** Engine may run alone; scheduler requires engine. Starting scheduler starts engine if needed. */
        function startService(id) {
            var s = getService(id);
            if (!s || s.running) return;
            if (id === "ebob-scheduler") {
                var eng = getService("ebob-engine");
                if (eng && !eng.running) {
                    eng.running = true;
                    applyStatusToRow("ebob-engine");
                }
            }
            s.running = true;
            applyStatusToRow(id);
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
        }

        function stopService(id) {
            var s = getService(id);
            if (!s || !s.running) return;
            if (id === "ebob-engine") {
                var sch = getService("ebob-scheduler");
                if (sch && sch.running) {
                    sch.running = false;
                    applyStatusToRow("ebob-scheduler");
                }
            }
            s.running = false;
            applyStatusToRow(id);
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
        }

        function restartService(id) {
            var s = getService(id);
            if (!s) return;
            if (s.running) {
                stopService(id);
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
                minimizeServicesMsc();
                return;
            }
            if (e.key === "e" || e.key === "E") {
                e.preventDefault();
                var id = EBOB_IDS[mscState.eBobCycle % EBOB_IDS.length];
                mscState.eBobCycle += 1;
                selectService(id);
                scrollRowIntoView(id);
                try {
                    var sel = window.getSelection && window.getSelection();
                    if (sel && sel.removeAllRanges) sel.removeAllRanges();
                } catch (err) {}
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

        (function wireSvcMscTableEvents() {
            var svcMscTableEl = getSvcMscTableEl();
            if (!svcMscTableEl) return;
            svcMscTableEl.addEventListener("click", function (e) {
                var tr = e.target.closest && e.target.closest("tr[data-service-id]");
                if (!tr) return;
                selectService(tr.getAttribute("data-service-id"));
            });
            svcMscTableEl.addEventListener("contextmenu", function (e) {
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
        })();

        document.addEventListener("keydown", onServicesKeydown, true);
        document.addEventListener("keydown", onGlobalEscape, true);

        document.addEventListener("click", function (e) {
            if (!svcCtxMenu || svcCtxMenu.hidden) return;
            if (e.target.closest && e.target.closest("#svcCtxMenu")) return;
            if (shouldSuppressSvcCtxMenuDismissForTutorial(e)) return;
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
        if (uacClose) {
            uacClose.addEventListener("click", function () {
                if (backdropUac) backdropUac.removeAttribute("data-uac-context");
                hideUac();
            });
        }
        if (uacMoreDetails && uacDetails) {
            uacMoreDetails.addEventListener("click", function () {
                uacDetails.hidden = !uacDetails.hidden;
                uacMoreDetails.setAttribute("aria-expanded", uacDetails.hidden ? "false" : "true");
                uacMoreDetails.textContent = uacDetails.hidden ? "Show more details" : "Hide details";
            });
        }

        if (svcMscClose) svcMscClose.addEventListener("click", hideServicesMsc);
        if (svcMscMin) svcMscMin.addEventListener("click", minimizeServicesMsc);
        if (svcMscMax && servicesMscShell) {
            svcMscMax.addEventListener("click", function () {
                servicesMscShell.classList.toggle("modal-services-msc--max");
            });
        }
        if (taskbarBtnServices) {
            taskbarBtnServices.addEventListener("click", function () {
                if (!backdropSvc) return;
                if (backdropSvc.classList.contains("show")) return;
                restoreServicesMsc();
            });
        }

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
            if (!state.ebobServicesRunning) state.ebobSchedulerRunning = false;
            var engSync = getService("ebob-engine");
            var schSync = getService("ebob-scheduler");
            if (engSync) engSync.running = !!state.ebobServicesRunning;
            if (schSync) {
                schSync.running = !!state.ebobSchedulerRunning && !!state.ebobServicesRunning;
            }
            renderServicesTable();
            updateDescAndToolbar();
            syncMscEngineToWorkstation();
            toast("Refreshed.");
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

        window.__ebobShowServicesUac = showUac;
    })();
})();
