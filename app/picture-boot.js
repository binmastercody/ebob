/* Runs first on lesson pages. A hidden iframe with ?picture=1 gets its own memory
   storage so building a picture guide never changes the user's saved progress. */
(function () {
    try {
        if (new URLSearchParams(location.search).get("picture") !== "1") return;
        function mem() {
            var s = Object.create(null);
            return {
                getItem: function (k) { return Object.prototype.hasOwnProperty.call(s, k) ? s[k] : null; },
                setItem: function (k, v) { s[k] = String(v); },
                removeItem: function (k) { delete s[k]; },
                clear: function () { s = Object.create(null); },
                key: function (i) { return Object.keys(s)[i] || null; },
                get length() { return Object.keys(s).length; }
            };
        }
        try {
            Object.defineProperty(window, "localStorage", { configurable: true, value: mem() });
            Object.defineProperty(window, "sessionStorage", { configurable: true, value: mem() });
        } catch (e) {
            try { localStorage.clear(); sessionStorage.clear(); } catch (e2) {}
        }
        document.documentElement.classList.add("eg-picture-mode");
    } catch (e) {}
})();
