// SendShield core logic: pure functions shared by the send handler, the task pane and the tests.
// Written in ES5 so it runs in every Outlook runtime (including the classic Outlook JS-only runtime).
(function (root) {
    "use strict";

    function normalizeDomain(value) {
        return String(value || "").trim().toLowerCase().replace(/^@/, "").replace(/\.$/, "");
    }

    function normalizeAddress(value) {
        return String(value || "").trim().toLowerCase();
    }

    function domainOf(address) {
        var email = normalizeAddress(address);
        var at = email.lastIndexOf("@");
        if (at < 1 || at === email.length - 1) return "";
        return normalizeDomain(email.substring(at + 1));
    }

    function endsWith(text, suffix) {
        return text.length >= suffix.length && text.substring(text.length - suffix.length) === suffix;
    }

    function domainInList(domain, list, includeSubdomains) {
        domain = normalizeDomain(domain);
        if (!domain || !list) return false;
        for (var i = 0; i < list.length; i++) {
            var entry = normalizeDomain(list[i]);
            if (!entry) continue;
            if (domain === entry) return true;
            if (includeSubdomains && endsWith(domain, "." + entry)) return true;
        }
        return false;
    }

    // Splits recipients into external / internal / ignored. Externals are de-duplicated and lower-cased.
    function classifyRecipients(recipients, config) {
        var result = { externals: [], externalDomains: [], internalCount: 0, ignoredCount: 0 };
        var seen = {};
        var seenDomain = {};
        var ignored = {};
        var includeSub = config.includeSubdomains !== false;
        var i;

        for (i = 0; i < (config.ignoredAddresses || []).length; i++) {
            ignored[normalizeAddress(config.ignoredAddresses[i])] = true;
        }

        for (i = 0; i < (recipients || []).length; i++) {
            var address = normalizeAddress(recipients[i] && recipients[i].emailAddress);
            var domain = domainOf(address);
            // Exchange-only addresses (no "@") resolve inside the organisation.
            if (!domain) { result.internalCount++; continue; }
            if (ignored[address]) { result.ignoredCount++; continue; }
            if (domainInList(domain, config.internalDomains, includeSub)) { result.internalCount++; continue; }
            if (domainInList(domain, config.trustedDomains, includeSub)) { result.ignoredCount++; continue; }
            if (seen[address]) continue;
            seen[address] = true;
            result.externals.push(address);
            if (!seenDomain[domain]) {
                seenDomain[domain] = true;
                result.externalDomains.push(domain);
            }
        }
        return result;
    }

    // Items in `items` that are not in `confirmed` (case-insensitive).
    function pending(items, confirmed) {
        var done = {};
        var out = [];
        var i;
        for (i = 0; i < (confirmed || []).length; i++) done[String(confirmed[i]).toLowerCase()] = true;
        for (i = 0; i < (items || []).length; i++) {
            if (!done[String(items[i]).toLowerCase()]) out.push(items[i]);
        }
        return out;
    }

    function union(a, b) {
        var seen = {};
        var out = [];
        var all = (a || []).concat(b || []);
        for (var i = 0; i < all.length; i++) {
            var key = String(all[i]).toLowerCase();
            if (seen[key]) continue;
            seen[key] = true;
            out.push(all[i]);
        }
        return out;
    }

    function parseList(raw) {
        if (!raw) return [];
        try {
            var value = JSON.parse(raw);
            return Object.prototype.toString.call(value) === "[object Array]" ? value : [];
        } catch (e) {
            return [];
        }
    }

    function formatList(items, max) {
        max = max || 3;
        if (items.length <= max) return items.join(", ");
        return items.slice(0, max).join(", ") + " (+" + (items.length - max) + " more)";
    }

    function truncate(text, limit) {
        text = String(text || "");
        if (text.length <= limit) return text;
        return text.substring(0, limit - 3) + "...";
    }

    function fillTemplate(template, values) {
        return String(template || "").replace(/\{(\w+)\}/g, function (match, key) {
            return Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : match;
        });
    }

    // Shallow merge with one level of nesting for `messages`.
    function mergeConfig(defaults, overrides) {
        var out = {};
        var key;
        for (key in defaults) if (Object.prototype.hasOwnProperty.call(defaults, key)) out[key] = defaults[key];
        if (overrides && typeof overrides === "object") {
            for (key in overrides) {
                if (!Object.prototype.hasOwnProperty.call(overrides, key)) continue;
                if (key === "messages") continue;
                if (overrides[key] !== undefined && overrides[key] !== null) out[key] = overrides[key];
            }
        }
        out.messages = {};
        var dm = defaults.messages || {};
        var om = (overrides && overrides.messages) || {};
        for (key in dm) if (Object.prototype.hasOwnProperty.call(dm, key)) out.messages[key] = dm[key];
        for (key in om) if (Object.prototype.hasOwnProperty.call(om, key) && om[key]) out.messages[key] = om[key];
        return out;
    }

    // Decides what the send handler should do. Pure, so it can be unit tested.
    //   input: { config, recipients, attachments, session, sessionAvailable }
    //   returns: { action: "allow" | "block" | "fallback", ... }
    function evaluateSend(input) {
        var config = input.config;
        if (!config.enabled) return { action: "allow", reason: "disabled" };

        var cls = classifyRecipients(input.recipients, config);
        if (cls.externals.length === 0) return { action: "allow", reason: "internal", classification: cls };

        var session = input.session || {};
        var attachments = config.confirmAttachments ? (input.attachments || []) : [];
        var pendingRecipients = pending(cls.externals, parseList(session.ssConfirmed));
        var pendingAttachments = pending(attachments, parseList(session.ssAttConfirmed));

        var base = {
            classification: cls,
            attachments: attachments,
            pendingRecipients: pendingRecipients,
            pendingAttachments: pendingAttachments
        };

        if (pendingRecipients.length === 0 && pendingAttachments.length === 0) {
            base.action = "allow";
            base.reason = "confirmed";
            return base;
        }

        var blocks = parseInt(session.ssBlocks || "0", 10);
        if (isNaN(blocks)) blocks = 0;
        base.blocks = blocks + 1;

        // Loop-breaker: if confirmations can't be stored, or the user has been blocked too often
        // for this draft, never trap them - fall back to a dialog that offers "Send anyway".
        if (!input.sessionAvailable || base.blocks > config.maxBlocksBeforeFallback) {
            base.action = "fallback";
            base.reason = input.sessionAvailable ? "loop_breaker" : "no_session";
            return base;
        }

        base.action = "block";
        base.reason = "unconfirmed";
        return base;
    }

    var api = {
        normalizeDomain: normalizeDomain,
        domainOf: domainOf,
        domainInList: domainInList,
        classifyRecipients: classifyRecipients,
        pending: pending,
        union: union,
        parseList: parseList,
        formatList: formatList,
        truncate: truncate,
        fillTemplate: fillTemplate,
        mergeConfig: mergeConfig,
        evaluateSend: evaluateSend
    };

    root.SendShieldCore = api;
})(typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : this);
