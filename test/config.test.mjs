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
