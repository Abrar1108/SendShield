import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { DEFAULT_CONFIG } from "../lib/defaults.js";

const context = {};
vm.createContext(context);
vm.runInContext(readFileSync(new URL("../src/addin/core.js", import.meta.url), "utf8"), context);
const Core = context.SendShieldCore;
const config = (overrides = {}) => Core.mergeConfig(DEFAULT_CONFIG, overrides);
const r = (...emails) => emails.map((emailAddress) => ({ emailAddress }));

test("internal, subdomain and Exchange-only addresses are not external", () => {
    const cls = Core.classifyRecipients(r("a@toluna.com", "b@uk.toluna.com", "c@MetrixLab.com", "/o=ExchangeLabs/cn=x"), config());
    assert.deepEqual([...cls.externals], []);
    assert.equal(cls.internalCount, 4);
});

test("lookalike and suffix domains are external", () => {
    const cls = Core.classifyRecipients(r("a@toluna.com.evil.io", "b@nottoluna.com", "c@tolunna.com"), config());
    assert.equal(cls.externals.length, 3);
});

test("subdomains can be excluded", () => {
    const cls = Core.classifyRecipients(r("b@uk.toluna.com"), config({ includeSubdomains: false }));
    assert.deepEqual([...cls.externals], ["b@uk.toluna.com"]);
});

test("ignored addresses and trusted domains are skipped, externals de-duplicated", () => {
    const cls = Core.classifyRecipients(
        r("X@Client.com", "x@client.com", "log@outreach.example", "p@partner.com"),
        config({ ignoredAddresses: ["log@outreach.example"], trustedDomains: ["partner.com"] })
    );
    assert.deepEqual([...cls.externals], ["x@client.com"]);
    assert.equal(cls.ignoredCount, 2);
});

test("internal-only email is allowed", () => {
    const d = Core.evaluateSend({ config: config(), recipients: r("a@toluna.com"), attachments: [], session: {}, sessionAvailable: true });
    assert.equal(d.action, "allow");
    assert.equal(d.reason, "internal");
});

test("disabled config always allows", () => {
    const d = Core.evaluateSend({ config: config({ enabled: false }), recipients: r("x@client.com"), session: {}, sessionAvailable: true });
    assert.equal(d.action, "allow");
});

test("unconfirmed external is blocked, then allowed once confirmed", () => {
    const base = { config: config(), recipients: r("a@toluna.com", "x@client.com"), attachments: [], sessionAvailable: true };
    assert.equal(Core.evaluateSend({ ...base, session: {} }).action, "block");
    const ok = Core.evaluateSend({ ...base, session: { ssConfirmed: JSON.stringify(["x@client.com"]) } });
    assert.equal(ok.action, "allow");
    assert.equal(ok.reason, "confirmed");
});

test("confirmation is per address: adding ignored or internal recipients does not re-prompt", () => {
    const session = { ssConfirmed: JSON.stringify(["x@client.com"]) };
    const d = Core.evaluateSend({
        config: config({ ignoredAddresses: ["log@outreach.example"] }),
        recipients: r("x@client.com", "log@outreach.example", "new@toluna.com"),
        attachments: [], session, sessionAvailable: true
    });
    assert.equal(d.action, "allow");
});

test("a newly added external is the only one pending", () => {
    const d = Core.evaluateSend({
        config: config(), recipients: r("x@client.com", "y@other.com"), attachments: [],
        session: { ssConfirmed: JSON.stringify(["x@client.com"]) }, sessionAvailable: true
    });
    assert.equal(d.action, "block");
    assert.deepEqual([...d.pendingRecipients], ["y@other.com"]);
});

test("attachments must be confirmed when enabled", () => {
    const base = { recipients: r("x@client.com"), attachments: ["report.xlsx"], sessionAvailable: true,
        session: { ssConfirmed: JSON.stringify(["x@client.com"]) } };
    assert.equal(Core.evaluateSend({ ...base, config: config() }).action, "block");
    assert.equal(Core.evaluateSend({ ...base, config: config({ confirmAttachments: false }) }).action, "allow");
    const ok = Core.evaluateSend({ ...base, config: config(), session: { ...base.session, ssAttConfirmed: JSON.stringify(["report.xlsx"]) } });
    assert.equal(ok.action, "allow");
});

test("loop-breaker kicks in after the configured number of blocks", () => {
    const base = { config: config({ maxBlocksBeforeFallback: 2 }), recipients: r("x@client.com"), attachments: [], sessionAvailable: true };
    assert.equal(Core.evaluateSend({ ...base, session: { ssBlocks: "0" } }).action, "block");
    assert.equal(Core.evaluateSend({ ...base, session: { ssBlocks: "1" } }).action, "block");
    const d = Core.evaluateSend({ ...base, session: { ssBlocks: "2" } });
    assert.equal(d.action, "fallback");
    assert.equal(d.reason, "loop_breaker");
});

test("without session storage the user gets the fallback dialog, never a dead end", () => {
    const d = Core.evaluateSend({ config: config(), recipients: r("x@client.com"), attachments: [], session: {}, sessionAvailable: false });
    assert.equal(d.action, "fallback");
    assert.equal(d.reason, "no_session");
});

test("corrupt session data is treated as unconfirmed", () => {
    const d = Core.evaluateSend({ config: config(), recipients: r("x@client.com"), attachments: [],
        session: { ssConfirmed: "{not json", ssBlocks: "abc" }, sessionAvailable: true });
    assert.equal(d.action, "block");
    assert.equal(d.blocks, 1);
});

test("formatting helpers", () => {
    assert.equal(Core.formatList(["a", "b", "c", "d", "e"], 3), "a, b, c (+2 more)");
    assert.equal(Core.truncate("x".repeat(300), 250).length, 250);
    assert.equal(Core.fillTemplate("{count} to {list} {missing}", { count: 2, list: "a, b" }), "2 to a, b {missing}");
});
