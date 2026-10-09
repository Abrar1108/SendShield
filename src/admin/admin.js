// SendShield admin dashboard. Plain JS, no framework, no inline handlers (CSP-safe).
(function () {
    "use strict";

    var LABELS = {
        prompt_shown: "Prompts shown",
        taskpane_confirmed: "Confirmed in task pane",
        sent_confirmed: "Sent externally",
        loop_breaker: "Loop-breaker",
        fail_open: "Fail-open"
    };
    var PLATFORM_LABELS = { PC: "Classic Outlook (Windows)", OfficeOnline: "Outlook on the web / new Outlook", Mac: "Outlook for Mac", iOS: "iOS", Android: "Android", unknown: "Unknown" };
    var MESSAGE_KEYS = ["blockPrompt", "attachmentPrompt", "fallbackPrompt", "reviewButton", "banner", "taskpaneIntro"];

    var state = { session: null, config: null, saved: null, days: 30, report: null };

    // ------------------------------------------------------------------ helpers
    function $(id) { return document.getElementById(id); }
    function qsa(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
    function el(tag, cls, text) {
        var node = document.createElement(tag);
        if (cls) node.className = cls;
        if (text !== undefined) node.textContent = text;
        return node;
    }
    function fmt(n) { return Number(n || 0).toLocaleString(); }
    function when(iso) {
        if (!iso) return "never";
        var d = new Date(iso);
        return d.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" }) + " " +
            d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
    }

    function api(method, url, body) {
        var options = { method: method, credentials: "same-origin", headers: { "X-SendShield": "1" } };
        if (body !== undefined) {
            options.headers["Content-Type"] = "application/json";
            options.body = JSON.stringify(body);
        }
        return fetch(url, options).then(function (res) {
            return res.text().then(function (text) {
                var data = {};
                try { data = text ? JSON.parse(text) : {}; } catch (e) { data = {}; }
                if (!res.ok) {
                    var err = new Error(data.error || ("Request failed (" + res.status + ")"));
                    err.status = res.status;
                    err.data = data;
                    throw err;
                }
                return data;
            });
        });
    }

    var toastTimer;
    function toast(text) {
        var t = $("toast");
        t.textContent = text;
        t.hidden = false;
        clearTimeout(toastTimer);
        toastTimer = setTimeout(function () { t.hidden = true; }, 3200);
    }

    // Fade-in on entry, staggered by data-index.
    var observer = "IntersectionObserver" in window ? new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
            if (entry.isIntersecting) { entry.target.classList.add("in"); observer.unobserve(entry.target); }
        });
    }, { threshold: 0.08 }) : null;
    function reveal(root) {
        qsa(".reveal:not(.in)", root).forEach(function (node) {
            var index = node.getAttribute("data-index");
            if (index) node.style.setProperty("--index", index);
            if (observer) observer.observe(node); else node.classList.add("in");
        });
    }

    // ------------------------------------------------------------------ session
    function showLogin(message) {
        $("app-view").hidden = true;
        $("user-box").hidden = true;
        $("login-view").hidden = false;
        if (message) { $("login-error").textContent = message; $("login-error").hidden = false; }
        reveal($("login-view"));
        setTimeout(function () { $("login-name").focus(); }, 50);
    }

    function showApp() {
        $("login-view").hidden = true;
        $("app-view").hidden = false;
        $("user-box").hidden = false;
        $("user-name").textContent = state.session.name;
        $("storage-warning").hidden = !!state.session.storage;
        selectTab(location.hash.replace("#", "") || "overview");
        reveal($("app-view"));
        loadConfig();
        loadReport();
    }

    function start() {
        api("GET", "/api/admin/session").then(function (s) {
            state.session = s;
            if (!s.passwordConfigured) {
                showLogin("ADMIN_PASSWORD is not set in Vercel (minimum 12 characters). Add it under Settings > Environment Variables, then redeploy.");
                $("login-submit").disabled = true;
                return;
            }
            if (s.signedIn) showApp(); else showLogin();
        }).catch(function () { showLogin("Can't reach the SendShield API."); });
    }

    $("login-form").addEventListener("submit", function (e) {
        e.preventDefault();
        $("login-error").hidden = true;
        $("login-submit").disabled = true;
        api("POST", "/api/admin/login", { name: $("login-name").value, password: $("login-password").value })
            .then(function () {
                $("login-password").value = "";
                return api("GET", "/api/admin/session");
            })
            .then(function (s) { state.session = s; showApp(); })
            .catch(function (err) {
                $("login-error").textContent = err.message;
                $("login-error").hidden = false;
            })
            .then(function () { $("login-submit").disabled = false; });
    });

    $("sign-out").addEventListener("click", function () {
        api("POST", "/api/admin/logout").then(function () { location.hash = ""; location.reload(); });
    });

    function handleAuthError(err) {
        if (err.status === 401) { showLogin("Your session has ended. Please sign in again."); return true; }
        return false;
    }

    // ------------------------------------------------------------------ tabs
    function selectTab(name) {
        var tabs = qsa(".tabs [data-tab]");
        if (!tabs.some(function (t) { return t.getAttribute("data-tab") === name; })) name = "overview";
        tabs.forEach(function (t) { t.setAttribute("aria-selected", String(t.getAttribute("data-tab") === name)); });
        qsa("[data-panel]").forEach(function (p) { p.hidden = p.getAttribute("data-panel") !== name; });
        if (location.hash.replace("#", "") !== name) history.replaceState(null, "", "#" + name);
        reveal(document.querySelector('[data-panel="' + name + '"]'));
        if (name === "audit") loadAudit();
        if (name === "overview" && state.report) drawChart(state.report.days);
    }
    qsa(".tabs [data-tab]").forEach(function (t) {
        t.addEventListener("click", function () { selectTab(t.getAttribute("data-tab")); });
    });

    // ------------------------------------------------------------------ header status
    function renderStatus(data) {
        var tag = $("protection-tag");
        var on = data.config.enabled;
        tag.textContent = on ? "Protection on" : "Protection paused";
        tag.className = "tag " + (on ? "green" : "red");
        $("meta-line").textContent = data.updatedAt
            ? "Settings last changed " + when(data.updatedAt) + (data.updatedBy ? " by " + data.updatedBy : "")
            : "Using built-in default settings";
    }

    // ------------------------------------------------------------------ settings
    var form = $("settings-form");

    function lines(list) { return (list || []).join("\n"); }
    function parseLines(text) { return text.split(/[\s,;]+/).map(function (s) { return s.trim(); }).filter(Boolean); }

    function fillForm(cfg) {
        form.elements.enabled.checked = !!cfg.enabled;
        form.elements.internalDomains.value = lines(cfg.internalDomains);
        form.elements.includeSubdomains.checked = !!cfg.includeSubdomains;
        form.elements.trustedDomains.value = lines(cfg.trustedDomains);
        form.elements.ignoredAddresses.value = lines(cfg.ignoredAddresses);
        form.elements.confirmAttachments.checked = !!cfg.confirmAttachments;
        form.elements.showRecipientBanner.checked = !!cfg.showRecipientBanner;
        form.elements.telemetryEnabled.checked = !!cfg.telemetryEnabled;
        form.elements.failOpenTimeoutMs.value = cfg.failOpenTimeoutMs;
        form.elements.maxBlocksBeforeFallback.value = cfg.maxBlocksBeforeFallback;
        MESSAGE_KEYS.forEach(function (k) { form.elements["messages." + k].value = (cfg.messages || {})[k] || ""; });
        updateCounters();
    }

    function readForm() {
        var messages = {};
        MESSAGE_KEYS.forEach(function (k) { messages[k] = form.elements["messages." + k].value; });
        return {
            enabled: form.elements.enabled.checked,
            internalDomains: parseLines(form.elements.internalDomains.value),
            includeSubdomains: form.elements.includeSubdomains.checked,
            trustedDomains: parseLines(form.elements.trustedDomains.value),
            ignoredAddresses: parseLines(form.elements.ignoredAddresses.value),
            confirmAttachments: form.elements.confirmAttachments.checked,
            showRecipientBanner: form.elements.showRecipientBanner.checked,
            telemetryEnabled: form.elements.telemetryEnabled.checked,
            failOpenTimeoutMs: Number(form.elements.failOpenTimeoutMs.value),
            maxBlocksBeforeFallback: Number(form.elements.maxBlocksBeforeFallback.value),
            messages: messages
        };
    }

    function updateCounters() {
        qsa(".counter").forEach(function (c) {
            var input = form.elements["messages." + c.getAttribute("data-for")];
            var max = Number(input.getAttribute("maxlength"));
            c.textContent = input.value.length + " / " + max;
            c.classList.toggle("over", input.value.length >= max);
        });
    }

    function updateDirty() {
        var dirty = state.saved && JSON.stringify(readForm()) !== state.saved;
        $("save").disabled = !dirty;
        $("discard").disabled = !dirty;
        $("savebar").classList.toggle("dirty", !!dirty);
        $("save-state").textContent = dirty ? "Unsaved changes" : "No unsaved changes";
    }

    form.addEventListener("input", function () { updateCounters(); updateDirty(); });
    form.addEventListener("change", updateDirty);

    form.elements.enabled.addEventListener("change", function (e) {
        if (!e.target.checked && !window.confirm("Pause SendShield for everyone? Users will be able to send external email without confirming once you save.")) {
            e.target.checked = true;
            updateDirty();
        }
    });

    $("discard").addEventListener("click", function () {
        fillForm(state.config);
        state.saved = JSON.stringify(readForm());
        $("settings-errors").hidden = true;
        updateDirty();
    });

    form.addEventListener("submit", function (e) {
        e.preventDefault();
        $("save").disabled = true;
        $("settings-errors").hidden = true;
        api("PUT", "/api/admin/config", { config: readForm() }).then(function (data) {
            state.config = data.config;
            fillForm(data.config);
            state.saved = JSON.stringify(readForm());
            renderStatus(data);
            updateDirty();
            toast(data.storage ? "Settings saved" : "Saved for this session only - storage is not connected");
        }).catch(function (err) {
            if (handleAuthError(err)) return;
            var box = $("settings-errors");
            box.textContent = err.message;
            if (err.data && err.data.errors) {
                var ul = el("ul");
                err.data.errors.forEach(function (msg) { ul.appendChild(el("li", "", msg)); });
                box.appendChild(ul);
            }
            box.hidden = false;
            updateDirty();
        });
    });

    function loadConfig() {
        api("GET", "/api/admin/config").then(function (data) {
            state.config = data.config;
            fillForm(data.config);
            state.saved = JSON.stringify(readForm());
            renderStatus(data);
            updateDirty();
        }).catch(function (err) { if (!handleAuthError(err)) toast(err.message); });
    }

    // ------------------------------------------------------------------ reports
    qsa("#range [data-days]").forEach(function (b) {
        b.addEventListener("click", function () {
            state.days = Number(b.getAttribute("data-days"));
            qsa("#range [data-days]").forEach(function (x) { x.setAttribute("aria-checked", String(x === b)); });
            loadReport();
        });
    });

    function loadReport() {
        api("GET", "/api/admin/reports?days=" + state.days).then(function (report) {
            state.report = report;
            renderReport(report);
        }).catch(function (err) { if (!handleAuthError(err)) toast(err.message); });
    }

    function renderReport(report) {
        var t = report.totals || {};
        ["prompt_shown", "taskpane_confirmed", "sent_confirmed"].forEach(function (k) {
            $("kpi-" + k).textContent = report.storage ? fmt(t[k]) : "–";
        });
        $("kpi-safety").textContent = report.storage ? fmt((t.loop_breaker || 0) + (t.fail_open || 0)) : "–";
        $("kpi-safety-note").textContent = report.storage
            ? fmt(t.loop_breaker) + " loop-breaker, " + fmt(t.fail_open) + " fail-open"
            : "Loop-breaker and fail-open events";

        drawChart(report.days || []);
        renderDailyTable(report.days || []);
        renderRank($("top-domains"), (report.topDomains || []).slice(0, 10).map(function (d) { return { name: d.domain, count: d.count }; }),
            report.storage ? "No external emails recorded yet." : "Connect storage to see domains.");
        var platforms = Object.keys(report.platforms || {}).map(function (p) {
            return { name: PLATFORM_LABELS[p] || p, count: report.platforms[p] };
        }).sort(function (a, b) { return b.count - a.count; });
        renderRank($("platforms"), platforms, report.storage ? "No events recorded yet." : "Connect storage to see clients.");
    }

    function renderRank(list, rows, emptyText) {
        list.innerHTML = "";
        if (!rows.length) { list.appendChild(el("li", "empty-row", emptyText)); return; }
        var max = rows[0].count || 1;
        rows.forEach(function (r) {
            var li = el("li");
            li.appendChild(el("span", "name", r.name));
            li.appendChild(el("span", "num", fmt(r.count)));
            var track = el("span", "bar-track");
            var fill = el("span", "bar-fill");
            fill.style.transform = "scaleX(" + Math.max(0.02, r.count / max) + ")";
            track.appendChild(fill);
            li.appendChild(track);
            list.appendChild(li);
        });
    }

    function renderDailyTable(days) {
        var wrap = $("daily-table");
        wrap.innerHTML = "";
        var table = el("table", "table");
        var head = el("tr");
        head.appendChild(el("th", "", "Date"));
        Object.keys(LABELS).forEach(function (k) { head.appendChild(el("th", "num", LABELS[k])); });
        var thead = el("thead"); thead.appendChild(head); table.appendChild(thead);
        var tbody = el("tbody");
        days.slice().reverse().forEach(function (d) {
            var tr = el("tr");
            tr.appendChild(el("td", "", d.date));
            Object.keys(LABELS).forEach(function (k) { tr.appendChild(el("td", "num", fmt(d[k]))); });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        wrap.appendChild(table);
    }

    $("toggle-table").addEventListener("click", function (e) {
        var open = $("daily-table").hidden;
        $("daily-table").hidden = !open;
        e.target.textContent = open ? "Hide table" : "Show table";
        e.target.setAttribute("aria-expanded", String(open));
    });

    // Single-series bar chart: prompts shown per day, hover tooltip shows every count for that day.
    var SVG = "http://www.w3.org/2000/svg";
    function svg(tag, attrs) {
        var node = document.createElementNS(SVG, tag);
        for (var k in attrs) if (Object.prototype.hasOwnProperty.call(attrs, k)) node.setAttribute(k, attrs[k]);
        return node;
    }
    function niceMax(v) {
        if (v <= 4) return 4;
        var p = Math.pow(10, Math.floor(Math.log10(v)));
        var steps = [1, 2, 2.5, 5, 10];
        for (var i = 0; i < steps.length; i++) if (steps[i] * p >= v) return steps[i] * p;
        return 10 * p;
    }
    function shortDate(iso) {
        var d = new Date(iso + "T00:00:00Z");
        return d.toLocaleDateString(undefined, { day: "numeric", month: "short", timeZone: "UTC" });
    }

    function drawChart(days) {
        var host = $("chart");
        host.innerHTML = "";
        var width = host.clientWidth || 800;
        var height = host.clientHeight || 240;
        if (!width) return;
        var m = { top: 8, right: 4, bottom: 24, left: 36 };
        var w = width - m.left - m.right;
        var h = height - m.top - m.bottom;
        var root = svg("svg", { viewBox: "0 0 " + width + " " + height, preserveAspectRatio: "none" });
        host.appendChild(root);

        var values = days.map(function (d) { return d.prompt_shown || 0; });
        var max = niceMax(Math.max.apply(null, values.concat([0])));
        var ticks = 4;
        for (var i = 0; i <= ticks; i++) {
            var val = (max / ticks) * i;
            var y = m.top + h - (val / max) * h;
            root.appendChild(svg("line", { "class": "grid", x1: m.left, x2: m.left + w, y1: y, y2: y }));
            var label = svg("text", { "class": "axis", x: m.left - 8, y: y + 4, "text-anchor": "end" });
            label.textContent = fmt(Math.round(val));
            root.appendChild(label);
        }

        if (!state.report || !state.report.storage || values.every(function (v) { return v === 0; })) {
            var empty = svg("text", { "class": "empty", x: m.left + w / 2, y: m.top + h / 2, "text-anchor": "middle" });
            empty.textContent = state.report && state.report.storage ? "No prompts in this period" : "Connect storage to see activity";
            root.appendChild(empty);
        }

        var n = days.length || 1;
        var slot = w / n;
        var gap = Math.max(2, slot * 0.25);
        var bw = Math.max(1, slot - gap);
        var labelEvery = Math.ceil(n / Math.max(2, Math.floor(w / 64)));
        var bars = [];

        days.forEach(function (d, idx) {
            var x = m.left + idx * slot + gap / 2;
            var bh = (values[idx] / max) * h;
            var y0 = m.top + h;
            if (bh > 0) {
                var r = Math.min(4, bw / 2, bh);
                var path = "M" + x + "," + y0 +
                    "V" + (y0 - bh + r) +
                    "Q" + x + "," + (y0 - bh) + " " + (x + r) + "," + (y0 - bh) +
                    "H" + (x + bw - r) +
                    "Q" + (x + bw) + "," + (y0 - bh) + " " + (x + bw) + "," + (y0 - bh + r) +
                    "V" + y0 + "Z";
                var bar = svg("path", { "class": "bar", d: path });
                root.appendChild(bar);
                bars[idx] = bar;
            }
            if (idx % labelEvery === 0 || idx === n - 1) {
                var t = svg("text", { "class": "axis", x: x + bw / 2, y: height - 6, "text-anchor": "middle" });
                t.textContent = shortDate(d.date);
                root.appendChild(t);
            }
            var hit = svg("rect", { "class": "hit", x: m.left + idx * slot, y: m.top, width: slot, height: h });
            hit.addEventListener("mouseenter", function () { showTip(idx); });
            hit.addEventListener("mouseleave", hideTip);
            root.appendChild(hit);
        });

        function showTip(idx) {
            host.classList.add("hovering");
            bars.forEach(function (b, j) { if (b) b.classList.toggle("active", j === idx); });
            var d = days[idx];
            var tip = $("tooltip");
            tip.innerHTML = "";
            tip.appendChild(el("strong", "", shortDate(d.date)));
            var dl = el("dl");
            Object.keys(LABELS).forEach(function (k) {
                dl.appendChild(el("dt", "", LABELS[k]));
                dl.appendChild(el("dd", "", fmt(d[k])));
            });
            tip.appendChild(dl);
            tip.hidden = false;
            var card = host.parentNode;
            var cx = host.offsetLeft + ((m.left + idx * slot + slot / 2) / width) * host.clientWidth;
            var left = cx + 12;
            if (left + tip.offsetWidth > card.clientWidth - 8) left = cx - tip.offsetWidth - 12;
            tip.style.left = Math.max(8, left) + "px";
            tip.style.top = (host.offsetTop + 8) + "px";
        }
        function hideTip() {
            host.classList.remove("hovering");
            bars.forEach(function (b) { if (b) b.classList.remove("active"); });
            $("tooltip").hidden = true;
        }
    }

    var resizeTimer;
    window.addEventListener("resize", function () {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(function () { if (state.report && !$("app-view").hidden) drawChart(state.report.days); }, 150);
    });

    // ------------------------------------------------------------------ audit
    function loadAudit() {
        api("GET", "/api/admin/audit").then(function (data) {
            var body = $("audit-body");
            body.innerHTML = "";
            if (!data.entries.length) {
                var tr = el("tr");
                var td = el("td", "muted", "No changes yet.");
                td.colSpan = 4;
                tr.appendChild(td);
                body.appendChild(tr);
                return;
            }
            data.entries.forEach(function (entry) {
                var tr = el("tr");
                tr.appendChild(el("td", "", when(entry.at)));
                tr.appendChild(el("td", "", entry.by || "–"));
                var changes = el("td");
                var chips = el("div", "chips");
                (entry.changes || []).forEach(function (c) { chips.appendChild(el("span", "chip", c)); });
                changes.appendChild(chips);
                tr.appendChild(changes);
                var status = el("td");
                status.appendChild(el("span", "tag " + (entry.enabled ? "green" : "red"), entry.enabled ? "On" : "Paused"));
                tr.appendChild(status);
                body.appendChild(tr);
            });
        }).catch(function (err) { if (!handleAuthError(err)) toast(err.message); });
    }

    // ------------------------------------------------------------------ install
    var manifestUrl = location.origin + "/manifest.xml";
    $("manifest-url").textContent = manifestUrl;
    $("copy-manifest").addEventListener("click", function () {
        if (navigator.clipboard) {
            navigator.clipboard.writeText(manifestUrl).then(function () { toast("Manifest URL copied"); }, function () { toast("Copy failed"); });
        }
    });

    start();
})();
