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

    var state = { session: null, config: null, saved: null, days: 30, report: null, meta: null };

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
        $("savebar").hidden = FORM_TABS.indexOf(name) === -1;
        if (name !== "messages" && name !== "settings") $("settings-errors").hidden = true;
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

    // ------------------------------------------------------------------ settings & messages
    var Core = window.SendShieldCore;
    var form = $("settings-form");
    var FORM_TABS = ["settings", "messages"];
    var GROUPS = {
        dialog: { title: "Send prompt", text: "Shown by Outlook when a user clicks Send with unconfirmed external recipients." },
        fallback: { title: "Safety-net prompt", text: "Shown instead after repeated blocked sends, with Outlook's Send anyway button." },
        banner: { title: "Compose banner", text: "Shown at the top of the email while external recipients are added." },
        taskpane: { title: "Task pane", text: "The panel where users confirm recipients and attachments." },
        ribbon: { title: "Outlook ribbon", text: "Stored in the manifest - after changing, download and upload the manifest again." }
    };
    var SAMPLE = {
        externals: ["anna.smith@unilever.com", "mark@nestle.com"],
        names: ["Anna Smith", "Mark Jones"],
        fields: ["to", "cc"],
        attachments: ["Q3 results.xlsx"],
        internal: 3
    };
    var previewMode = "dialog";

    function lines(list) { return (list || []).join("\n"); }
    function parseLines(text) { return text.split(/[\s,;]+/).map(function (s) { return s.trim(); }).filter(Boolean); }

    function buildMessageEditor(fields) {
        var host = $("message-groups");
        if (host.getAttribute("data-built")) return;
        host.setAttribute("data-built", "1");
        Object.keys(GROUPS).forEach(function (where, gi) {
            var card = el("article", "card msg-group reveal");
            card.setAttribute("data-index", String(gi));
            var head = el("div", "card-head");
            var titles = el("div");
            titles.appendChild(el("h3", "", GROUPS[where].title));
            titles.appendChild(el("p", "", GROUPS[where].text));
            head.appendChild(titles);
            card.appendChild(head);

            fields.filter(function (f) { return f.where === where; }).forEach(function (f) {
                var wrap = el("div", "msg-field");
                var label = el("label", "field");
                var top = el("span");
                top.appendChild(document.createTextNode(f.label));
                var dot = el("span", "modified-dot");
                dot.hidden = true;
                dot.title = "Changed from default";
                dot.setAttribute("data-dot", f.key);
                top.appendChild(dot);
                var counter = el("em", "counter");
                counter.setAttribute("data-for", f.key);
                top.appendChild(counter);
                label.appendChild(top);
                var input = f.multiline ? el("textarea") : el("input");
                if (!f.multiline) input.type = "text";
                else input.rows = 2;
                input.name = "messages." + f.key;
                input.maxLength = f.limit;
                input.setAttribute("data-where", f.where);
                input.addEventListener("focus", function () { setPreview(f.where === "fallback" ? "fallback" : f.where); });
                label.appendChild(input);
                wrap.appendChild(label);

                var tools = el("div", "msg-tools");
                f.vars.forEach(function (v) {
                    var chip = el("button", "var-chip", "{" + v + "}");
                    chip.type = "button";
                    chip.title = "Insert {" + v + "}";
                    chip.addEventListener("click", function () { insertAtCursor(input, "{" + v + "}"); });
                    tools.appendChild(chip);
                });
                var reset = el("button", "link-btn", "Reset to default");
                reset.type = "button";
                reset.setAttribute("data-reset", f.key);
                reset.addEventListener("click", function () {
                    input.value = state.meta.defaults.messages[f.key];
                    input.dispatchEvent(new Event("input", { bubbles: true }));
                });
                tools.appendChild(reset);
                wrap.appendChild(tools);
                card.appendChild(wrap);
            });
            host.appendChild(card);
        });
        reveal(host);
    }

    function insertAtCursor(input, text) {
        var start = input.selectionStart == null ? input.value.length : input.selectionStart;
        var end = input.selectionEnd == null ? input.value.length : input.selectionEnd;
        input.value = (input.value.slice(0, start) + text + input.value.slice(end)).slice(0, input.maxLength);
        input.focus();
        input.setSelectionRange(start + text.length, start + text.length);
        input.dispatchEvent(new Event("input", { bubbles: true }));
    }

    function messageKeys() { return (state.meta ? state.meta.messageFields : []).map(function (f) { return f.key; }); }

    function fillForm(cfg) {
        var e = form.elements;
        e.enabled.checked = !!cfg.enabled;
        e.pilotMode.checked = !!cfg.pilotMode;
        e.pilotUsers.value = lines(cfg.pilotUsers);
        e.internalDomains.value = lines(cfg.internalDomains);
        e.includeSubdomains.checked = !!cfg.includeSubdomains;
        e.trustedDomains.value = lines(cfg.trustedDomains);
        e.ignoredAddresses.value = lines(cfg.ignoredAddresses);
        e.confirmStyle.value = cfg.confirmStyle || "each";
        e.confirmAttachments.checked = !!cfg.confirmAttachments;
        e.showRecipientBanner.checked = !!cfg.showRecipientBanner;
        e.minExternalRecipients.value = cfg.minExternalRecipients;
        e.failOpenTimeoutMs.value = cfg.failOpenTimeoutMs;
        e.maxBlocksBeforeFallback.value = cfg.maxBlocksBeforeFallback;
        e.configRefreshMinutes.value = cfg.configRefreshMinutes;
        e.telemetryEnabled.checked = !!cfg.telemetryEnabled;
        messageKeys().forEach(function (k) { e["messages." + k].value = (cfg.messages || {})[k] || ""; });
        updateMessageUi();
    }

    function readForm() {
        var e = form.elements;
        var messages = {};
        messageKeys().forEach(function (k) { messages[k] = e["messages." + k].value; });
        return {
            enabled: e.enabled.checked,
            pilotMode: e.pilotMode.checked,
            pilotUsers: parseLines(e.pilotUsers.value),
            internalDomains: parseLines(e.internalDomains.value),
            includeSubdomains: e.includeSubdomains.checked,
            trustedDomains: parseLines(e.trustedDomains.value),
            ignoredAddresses: parseLines(e.ignoredAddresses.value),
            confirmStyle: e.confirmStyle.value,
            confirmAttachments: e.confirmAttachments.checked,
            showRecipientBanner: e.showRecipientBanner.checked,
            minExternalRecipients: Number(e.minExternalRecipients.value),
            failOpenTimeoutMs: Number(e.failOpenTimeoutMs.value),
            maxBlocksBeforeFallback: Number(e.maxBlocksBeforeFallback.value),
            configRefreshMinutes: Number(e.configRefreshMinutes.value),
            telemetryEnabled: e.telemetryEnabled.checked,
            messages: messages
        };
    }

    function updateMessageUi() {
        $("pilot-users-field").hidden = !form.elements.pilotMode.checked;
        qsa(".counter").forEach(function (c) {
            var input = form.elements["messages." + c.getAttribute("data-for")];
            if (!input) return;
            c.textContent = input.value.length + " / " + input.maxLength;
            c.classList.toggle("over", input.value.length >= input.maxLength);
        });
        if (state.meta) {
            qsa("[data-dot]").forEach(function (dot) {
                var key = dot.getAttribute("data-dot");
                var changed = form.elements["messages." + key].value !== state.meta.defaults.messages[key];
                dot.hidden = !changed;
                var reset = document.querySelector('[data-reset="' + key + '"]');
                if (reset) reset.hidden = !changed;
            });
        }
        renderPreview();
    }

    function updateDirty() {
        var dirty = state.saved && JSON.stringify(readForm()) !== state.saved;
        $("save").disabled = !dirty;
        $("discard").disabled = !dirty;
        $("savebar").classList.toggle("dirty", !!dirty);
        $("save-state").textContent = dirty ? "Unsaved changes" : "No unsaved changes";
    }

    form.addEventListener("input", function () { updateMessageUi(); updateDirty(); });
    form.addEventListener("change", function () { updateMessageUi(); updateDirty(); });

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

    window.addEventListener("beforeunload", function (e) {
        if (!$("save").disabled) { e.preventDefault(); e.returnValue = ""; }
    });

    form.addEventListener("submit", function (e) {
        e.preventDefault();
        $("save").disabled = true;
        $("settings-errors").hidden = true;
        api("PUT", "/api/admin/config", { config: readForm() }).then(function (data) {
            applyConfig(data);
            if (data.manifestChanged) {
                toast("Saved. Ribbon text changed - upload the new manifest from Install.");
            } else {
                toast(data.storage ? "Changes saved" : "Saved for this session only - storage is not connected");
            }
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
            box.scrollIntoView({ behavior: "smooth", block: "center" });
            updateDirty();
        });
    });

    function applyConfig(data) {
        state.meta = { defaults: data.defaults, messageFields: data.messageFields };
        buildMessageEditor(data.messageFields);
        state.config = data.config;
        fillForm(data.config);
        state.saved = JSON.stringify(readForm());
        renderStatus(data);
        updateDirty();
        $("manifest-version").textContent = data.manifestVersion;
        if (data.manifestChanged) $("manifest-notice").hidden = false;
    }

    function loadConfig() {
        api("GET", "/api/admin/config").then(applyConfig)
            .catch(function (err) { if (!handleAuthError(err)) toast(err.message); });
    }

    // ------------------------------------------------------------------ live preview
    qsa("#preview-switch [data-preview]").forEach(function (b) {
        b.addEventListener("click", function () { setPreview(b.getAttribute("data-preview")); });
    });

    function setPreview(mode) {
        previewMode = mode;
        qsa("#preview-switch [data-preview]").forEach(function (b) {
            b.setAttribute("aria-checked", String(b.getAttribute("data-preview") === mode));
        });
        renderPreview();
    }

    function m(key, values) {
        var input = form.elements["messages." + key];
        return Core.fillTemplate(input ? input.value : "", values || {});
    }

    function icon(size) {
        var img = el("img");
        img.src = "/addin/assets/sendshield-" + (size || 32) + ".png";
        img.alt = "";
        return img;
    }

    function olButton(text, primary) { return el("span", "ol-btn" + (primary ? " primary" : ""), text); }

    function renderPreview() {
        var host = $("preview");
        if (!host || !state.meta) return;
        host.innerHTML = "";
        var list = Core.formatList(SAMPLE.externals, 3);
        var count = SAMPLE.externals.length;

        if (previewMode === "dialog" || previewMode === "fallback") {
            var text, buttons;
            if (previewMode === "dialog") {
                var hint = m("hintModern", { button: form.elements["messages.reviewButton"].value });
                text = m("blockPrompt", { count: count, list: list, hint: hint });
                buttons = [olButton(form.elements["messages.reviewButton"].value, true)];
            } else {
                text = m("fallbackPrompt", { count: count, list: list });
                buttons = [olButton("Send Anyway", true), olButton("Don't Send")];
            }
            var dlg = el("div", "ol ol-dialog");
            var h = el("h4");
            h.appendChild(icon(32));
            h.appendChild(document.createTextNode(form.elements["messages.taskpaneTitle"].value || "SendShield"));
            dlg.appendChild(h);
            dlg.appendChild(el("p", "", Core.truncate(text, 250)));
            var row = el("div", "ol-buttons");
            buttons.forEach(function (b) { row.appendChild(b); });
            dlg.appendChild(row);
            host.appendChild(dlg);
            var c = el("p", "ol-count" + (text.length > 250 ? " over" : ""), text.length + " / 250 characters" + (text.length > 250 ? " - will be cut off" : ""));
            host.appendChild(c);
            if (previewMode === "dialog") {
                var legacy = el("p", "fine", "Older Outlook: " + m("hintLegacy", { ribbon: form.elements["messages.ribbonButton"].value }));
                host.appendChild(legacy);
            }
            return;
        }

        if (previewMode === "banner") {
            var compose = el("div", "ol ol-compose");
            var banner = el("div", "ol-banner");
            banner.appendChild(icon(16));
            banner.appendChild(el("span", "", Core.truncate(m("banner", { count: count, list: Core.formatList(SAMPLE.externals, 2) }), 150)));
            compose.appendChild(banner);
            compose.appendChild(el("div", "ol-compose-row", "To: Anna Smith; Mark Jones"));
            compose.appendChild(el("div", "ol-compose-row", "Subject: Q3 results"));
            compose.appendChild(el("div", "ol-compose-body"));
            host.appendChild(compose);
            return;
        }

        if (previewMode === "ribbon") {
            var ribbon = el("div", "ol ol-ribbon");
            var btn = el("div", "ol-ribbon-btn");
            btn.appendChild(icon(80));
            btn.appendChild(el("span", "", form.elements["messages.ribbonButton"].value));
            ribbon.appendChild(btn);
            ribbon.appendChild(el("div", "ol-ribbon-group", form.elements["messages.ribbonGroup"].value));
            host.appendChild(ribbon);
            var tip = el("div", "ol ol-tooltip");
            tip.appendChild(el("strong", "", form.elements["messages.ribbonButton"].value));
            tip.appendChild(el("span", "", form.elements["messages.ribbonTooltip"].value));
            host.appendChild(tip);
            return;
        }

        // Task pane
        var each = form.elements.confirmStyle.value !== "all";
        var pane = el("div", "ol ol-pane");
        var ph = el("div", "ol-pane-head");
        ph.appendChild(icon(32));
        ph.appendChild(el("span", "", m("taskpaneTitle")));
        pane.appendChild(ph);
        var body = el("div", "ol-pane-body");
        body.appendChild(el("div", "", m("taskpaneIntro")));
        var h5 = el("h5", "", m("recipientsHeading") + " (" + count + ")");
        if (each) h5.appendChild(el("span", "", m("selectAll")));
        body.appendChild(h5);
        var ul = el("div", "ol-list");
        SAMPLE.externals.forEach(function (addr, i) {
            var row = el("div");
            if (each) row.appendChild(el("span", "ol-box" + (i === 0 ? " on" : "")));
            row.appendChild(el("span", "", addr));
            ul.appendChild(row);
        });
        body.appendChild(ul);
        if (!each) {
            var all = el("div", "ol-check");
            all.appendChild(el("span", "ol-box"));
            all.appendChild(el("span", "", m("confirmAllCheckbox", { count: count })));
            body.appendChild(all);
        }
        if (form.elements.confirmAttachments.checked) {
            body.appendChild(el("h5", "", m("attachmentsHeading") + " (1)"));
            var al = el("div", "ol-list");
            al.appendChild(el("div", "", SAMPLE.attachments[0]));
            body.appendChild(al);
            var ac = el("div", "ol-check");
            ac.appendChild(el("span", "ol-box"));
            ac.appendChild(el("span", "", m("attachmentsCheckbox")));
            body.appendChild(ac);
        }
        body.appendChild(el("div", "ol-confirm", m("confirmButton")));
        body.appendChild(el("div", "ol-note", each ? m("remainingRecipients", { count: 1 }) : m("remainingAttachments")));
        body.appendChild(el("div", "ol-ok", m("confirmedMessage")));
        pane.appendChild(body);
        var foot = el("div", "ol-foot");
        foot.appendChild(el("span", "", m("internalSummary", { count: SAMPLE.internal })));
        foot.appendChild(el("span", "", "v6"));
        pane.appendChild(foot);
        host.appendChild(pane);
        var other = el("p", "fine", "Other task pane messages: “" + m("noExternal") + "”, “" + m("pilotInactive") + "”, “" + m("legacyNotice") + "”, “" + m("saveFailed") + "”");
        host.appendChild(other);
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
