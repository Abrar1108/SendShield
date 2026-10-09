// SendShield runtime helpers: cached settings, fire-and-forget telemetry and Outlook API wrappers.
// Placeholders (__BASE_URL__ etc.) are filled in by scripts/build.mjs at deploy time.
(function (root) {
    "use strict";

    var Core = root.SendShieldCore;
    var BASE_URL = "__BASE_URL__";
    var VERSION = "__VERSION__";
    var DEFAULTS = __DEFAULT_CONFIG__;
    var CACHE_KEY = "sendshield.config.v1";
    var REFRESH_AFTER_MS = 15 * 60 * 1000;

    var memoryConfig = null;
    var fetchedAt = 0;
    var refreshing = false;

    function hasServer() {
        return BASE_URL.indexOf("http") === 0;
    }

    // ---------------------------------------------------------------- storage
    function storageGet(key, callback) {
        try {
            if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
                OfficeRuntime.storage.getItem(key).then(function (v) { callback(v); }, function () { callback(null); });
                return;
            }
        } catch (e) { /* fall through */ }
        try { callback(root.localStorage ? root.localStorage.getItem(key) : null); } catch (e2) { callback(null); }
    }

    function storageSet(key, value) {
        try {
            if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
                OfficeRuntime.storage.setItem(key, value).then(function () {}, function () {});
                return;
            }
        } catch (e) { /* fall through */ }
        try { if (root.localStorage) root.localStorage.setItem(key, value); } catch (e2) { /* ignore */ }
    }

    // ---------------------------------------------------------------- http
    function httpRequest(method, url, body, timeoutMs, callback) {
        var finished = false;
        function done(err, text) {
            if (finished) return;
            finished = true;
            if (callback) callback(err, text);
        }
        var timer = setTimeout(function () { done(new Error("timeout")); }, timeoutMs);
        try {
            if (typeof fetch === "function") {
                var options = { method: method, cache: "no-store" };
                if (body !== undefined) {
                    // text/plain avoids a CORS pre-flight request.
                    options.headers = { "Content-Type": "text/plain" };
                    options.body = body;
                }
                fetch(url, options).then(function (response) {
                    return response.text().then(function (text) {
                        clearTimeout(timer);
                        if (response.ok) done(null, text); else done(new Error("HTTP " + response.status));
                    });
                }, function (err) {
                    clearTimeout(timer);
                    done(err || new Error("network"));
                });
                return;
            }
            if (typeof XMLHttpRequest !== "undefined") {
                var xhr = new XMLHttpRequest();
                xhr.open(method, url, true);
                if (body !== undefined) xhr.setRequestHeader("Content-Type", "text/plain");
                xhr.onreadystatechange = function () {
                    if (xhr.readyState !== 4) return;
                    clearTimeout(timer);
                    if (xhr.status >= 200 && xhr.status < 300) done(null, xhr.responseText); else done(new Error("HTTP " + xhr.status));
                };
                xhr.send(body);
                return;
            }
        } catch (e) {
            clearTimeout(timer);
            done(e);
            return;
        }
        clearTimeout(timer);
        done(new Error("no http"));
    }

    // ---------------------------------------------------------------- config
    function defaults() {
        return Core.mergeConfig(DEFAULTS, null);
    }

    // Never touches the network: memory -> local cache -> built-in defaults.
    function loadConfig(callback) {
        if (memoryConfig) { callback(memoryConfig); return; }
        storageGet(CACHE_KEY, function (raw) {
            var cached = null;
            try { cached = raw ? JSON.parse(raw) : null; } catch (e) { cached = null; }
            if (cached && cached.config) {
                memoryConfig = Core.mergeConfig(DEFAULTS, cached.config);
                fetchedAt = cached.fetchedAt || 0;
                callback(memoryConfig);
            } else {
                callback(defaults());
            }
        });
    }

    // Background refresh from /api/config. Never awaited on the send path.
    function refreshConfig(force, callback) {
        if (!hasServer()) { if (callback) callback(false); return; }
        if (refreshing || (!force && fetchedAt && Date.now() - fetchedAt < REFRESH_AFTER_MS)) {
            if (callback) callback(false);
            return;
        }
        refreshing = true;
        httpRequest("GET", BASE_URL + "/api/config", undefined, 4000, function (err, text) {
            refreshing = false;
            if (err) { if (callback) callback(false); return; }
            try {
                var payload = JSON.parse(text);
                if (payload && payload.config) {
                    memoryConfig = Core.mergeConfig(DEFAULTS, payload.config);
                    fetchedAt = Date.now();
                    storageSet(CACHE_KEY, JSON.stringify({ config: payload.config, fetchedAt: fetchedAt }));
                    if (callback) callback(true);
                    return;
                }
            } catch (e) { /* ignore */ }
            if (callback) callback(false);
        });
    }

    // ---------------------------------------------------------------- telemetry
    function clientInfo() {
        try {
            var d = Office.context.diagnostics;
            return { platform: String(d.platform || ""), host: String(d.version || "") };
        } catch (e) {
            return { platform: "", host: "" };
        }
    }

    // Only counts and domains are sent - never addresses, subjects or file names.
    function sendEvent(config, type, details) {
        if (!hasServer() || !config || !config.telemetryEnabled) return;
        details = details || {};
        var info = clientInfo();
        var payload = {
            type: type,
            externalCount: details.externalCount || 0,
            externalDomains: (details.externalDomains || []).slice(0, 20),
            attachmentCount: details.attachmentCount || 0,
            reason: details.reason || "",
            platform: info.platform,
            version: VERSION
        };
        httpRequest("POST", BASE_URL + "/api/event", JSON.stringify(payload), 5000, null);
    }

    // ---------------------------------------------------------------- Outlook wrappers
    function getRecipients(item, callback) {
        var fields = ["to", "cc", "bcc"];
        var all = [];
        var remaining = fields.length;
        var ok = true;
        function collect(field) {
            try {
                item[field].getAsync(function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        var list = result.value || [];
                        for (var i = 0; i < list.length; i++) {
                            all.push({ emailAddress: list[i].emailAddress, displayName: list[i].displayName, field: field });
                        }
                    } else {
                        ok = false;
                    }
                    if (--remaining === 0) callback(all, ok);
                });
            } catch (e) {
                ok = false;
                if (--remaining === 0) callback(all, ok);
            }
        }
        for (var f = 0; f < fields.length; f++) collect(fields[f]);
    }

    function getAttachmentNames(item, callback) {
        if (typeof item.getAttachmentsAsync !== "function") { callback([]); return; }
        try {
            item.getAttachmentsAsync(function (result) {
                var names = [];
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var atts = result.value || [];
                    for (var i = 0; i < atts.length; i++) {
                        if (!atts[i].isInline) names.push(atts[i].name);
                    }
                }
                callback(names);
            });
        } catch (e) {
            callback([]);
        }
    }

    function sessionAvailable(item) {
        return !!(item && item.sessionData && typeof item.sessionData.getAllAsync === "function");
    }

    function getSession(item, callback) {
        if (!sessionAvailable(item)) { callback({}, false); return; }
        try {
            item.sessionData.getAllAsync(function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) callback(result.value || {}, true);
                else callback({}, false);
            });
        } catch (e) {
            callback({}, false);
        }
    }

    function setSession(item, values, callback) {
        if (!sessionAvailable(item)) { if (callback) callback(false); return; }
        var keys = [];
        for (var k in values) if (Object.prototype.hasOwnProperty.call(values, k)) keys.push(k);
        var remaining = keys.length;
        var ok = true;
        if (remaining === 0) { if (callback) callback(true); return; }
        for (var i = 0; i < keys.length; i++) {
            (function (key) {
                try {
                    item.sessionData.setAsync(key, String(values[key]), function (result) {
                        if (result.status !== Office.AsyncResultStatus.Succeeded) ok = false;
                        if (--remaining === 0 && callback) callback(ok);
                    });
                } catch (e) {
                    ok = false;
                    if (--remaining === 0 && callback) callback(ok);
                }
            })(keys[i]);
        }
    }

    function supports(version) {
        try {
            return Office.context.requirements.isSetSupported("Mailbox", version);
        } catch (e) {
            return false;
        }
    }

    root.SendShieldRuntime = {
        BASE_URL: BASE_URL,
        VERSION: VERSION,
        loadConfig: loadConfig,
        refreshConfig: refreshConfig,
        sendEvent: sendEvent,
        getRecipients: getRecipients,
        getAttachmentNames: getAttachmentNames,
        getSession: getSession,
        setSession: setSession,
        sessionAvailable: sessionAvailable,
        supports: supports
    };
})(typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : this);
