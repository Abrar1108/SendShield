import { test } from "node:test";
import assert from "node:assert/strict";
import { sanitizeConfig } from "../lib/config.js";
import { validateEvent } from "../lib/stats.js";

test("defaults are valid", () => {
    const { config, errors } = sanitizeConfig({});
    assert.deepEqual(errors, []);
    assert.deepEqual(config.internalDomains, ["toluna.com", "metrixlab.com"]);
});

test("domains are normalised and invalid entries reported", () => {
    const { config, errors } = sanitizeConfig({ internalDomains: " Toluna.com\n@MetrixLab.com.\nnot a domain\ntoluna.com", trustedDomains: [] });
    assert.deepEqual(config.internalDomains, ["toluna.com", "metrixlab.com"]);
    assert.equal(errors.length, 3); // "not", "a" and "domain" each fail
});

test("empty internal domains is rejected", () => {
    const { errors } = sanitizeConfig({ internalDomains: [] });
    assert.ok(errors.some((e) => e.includes("at least one")));
});

test("numbers are clamped and messages limited", () => {
    const { config, errors } = sanitizeConfig({ failOpenTimeoutMs: 99999, maxBlocksBeforeFallback: 0, messages: { reviewButton: "x".repeat(30) } });
    assert.equal(config.failOpenTimeoutMs, 8000);
    assert.equal(config.maxBlocksBeforeFallback, 1);
    assert.equal(config.messages.reviewButton.length, 20);
    assert.ok(errors.length >= 1);
});

test("ignored addresses must be email addresses", () => {
    const { config, errors } = sanitizeConfig({ ignoredAddresses: ["Log@Outreach.io", "outreach.io"] });
    assert.deepEqual(config.ignoredAddresses, ["log@outreach.io"]);
    assert.equal(errors.length, 1);
});

test("telemetry events keep only types, valid domains and a clean platform", () => {
    assert.equal(validateEvent({ type: "nope" }), null);
    const e = validateEvent({ type: "sent_confirmed", externalDomains: ["Client.com", "bad domain", "client.com"], platform: "PC<script>" });
    assert.deepEqual(e, { type: "sent_confirmed", domains: ["client.com"], platform: "PCscript" });
});

import { hashString, publicConfig } from "../lib/config.js";
import { renderManifest, manifestVersion } from "../lib/manifest.js";
import { DEFAULT_CONFIG, MESSAGE_FIELDS } from "../lib/defaults.js";
import { readFileSync } from "node:fs";
import vm from "node:vm";

test("server and add-in hash pilot mailboxes identically; public config hides addresses", () => {
    const ctx = {};
    vm.createContext(ctx);
    vm.runInContext(readFileSync(new URL("../src/addin/core.js", import.meta.url), "utf8"), ctx);
    for (const email of ["pilot@toluna.com", "a.very.long.name+tag@metrixlab.com"]) {
        assert.equal(hashString(email), ctx.SendShieldCore.hashString(email));
    }
    const { config } = sanitizeConfig({ pilotMode: true, pilotUsers: ["Pilot@Toluna.com"] });
    const pub = publicConfig(config);
    assert.equal(pub.pilotUsers, undefined);
    assert.deepEqual(pub.pilotUserHashes, [hashString("pilot@toluna.com")]);
});

test("pilot mode needs at least one mailbox; ignore list accepts @domain", () => {
    assert.ok(sanitizeConfig({ pilotMode: true, pilotUsers: [] }).errors.length);
    const { config, errors } = sanitizeConfig({ ignoredAddresses: ["@Tracker.example", "bot@x.io"] });
    assert.deepEqual(errors, []);
    assert.deepEqual(config.ignoredAddresses, ["@tracker.example", "bot@x.io"]);
});

test("every message has a default within its limit", () => {
    for (const f of MESSAGE_FIELDS) {
        assert.ok(f.value.length <= f.limit, f.key);
        assert.equal(DEFAULT_CONFIG.messages[f.key], f.value);
    }
});

test("manifest escapes ribbon text and carries the revision", () => {
    const xml = renderManifest({ baseUrl: "https://x.vercel.app", revision: 3,
        messages: { ...DEFAULT_CONFIG.messages, ribbonButton: 'Check <"&">' } });
    assert.ok(xml.includes("Check &lt;&quot;&amp;&quot;&gt;"));
    assert.ok(xml.includes(`<Version>${manifestVersion(3)}</Version>`));
    assert.ok(!xml.includes("undefined"));
});
