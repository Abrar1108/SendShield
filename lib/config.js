import { DEFAULT_CONFIG, MESSAGE_LIMITS } from "./defaults.js";
import { kv } from "./store.js";

const CONFIG_KEY = "ss:config";
const AUDIT_KEY = "ss:audit";
const MAX_LIST = 500;
const DOMAIN_RE = /^(?=.{3,253}$)([a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/;
const ADDRESS_RE = /^[^\s@]{1,64}@([a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/;

function toList(value) {
    if (Array.isArray(value)) return value;
    if (typeof value === "string") return value.split(/[\s,;]+/);
    return [];
}

function cleanList(value, pattern, label, errors, normalize) {
    const out = [];
    const seen = new Set();
    for (const raw of toList(value)) {
        const item = normalize(String(raw));
        if (!item) continue;
        if (!pattern.test(item)) {
            errors.push(`${label}: "${item}" is not valid`);
            continue;
        }
        if (seen.has(item)) continue;
        seen.add(item);
        out.push(item);
    }
    if (out.length > MAX_LIST) errors.push(`${label}: at most ${MAX_LIST} entries`);
    return out.slice(0, MAX_LIST);
}

function clampInt(value, min, max, fallback) {
    const n = Number.parseInt(value, 10);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
}

const normDomain = (s) => s.trim().toLowerCase().replace(/^@/, "").replace(/\.$/, "");
const normAddress = (s) => s.trim().toLowerCase();

// Validates admin input. Returns { config, errors }; config is always complete and safe to store.
export function sanitizeConfig(input) {
    const src = input && typeof input === "object" ? input : {};
    const errors = [];
    const d = DEFAULT_CONFIG;

    const config = {
        enabled: src.enabled === undefined ? d.enabled : Boolean(src.enabled),
        internalDomains: cleanList(src.internalDomains ?? d.internalDomains, DOMAIN_RE, "Internal domains", errors, normDomain),
        includeSubdomains: src.includeSubdomains === undefined ? d.includeSubdomains : Boolean(src.includeSubdomains),
        trustedDomains: cleanList(src.trustedDomains ?? d.trustedDomains, DOMAIN_RE, "Trusted domains", errors, normDomain),
        ignoredAddresses: cleanList(src.ignoredAddresses ?? d.ignoredAddresses, ADDRESS_RE, "Ignored addresses", errors, normAddress),
        confirmAttachments: src.confirmAttachments === undefined ? d.confirmAttachments : Boolean(src.confirmAttachments),
        showRecipientBanner: src.showRecipientBanner === undefined ? d.showRecipientBanner : Boolean(src.showRecipientBanner),
        failOpenTimeoutMs: clampInt(src.failOpenTimeoutMs, 1000, 8000, d.failOpenTimeoutMs),
        maxBlocksBeforeFallback: clampInt(src.maxBlocksBeforeFallback, 1, 5, d.maxBlocksBeforeFallback),
        telemetryEnabled: src.telemetryEnabled === undefined ? d.telemetryEnabled : Boolean(src.telemetryEnabled),
        messages: {}
    };

    if (config.internalDomains.length === 0) {
        errors.push("Internal domains: at least one domain is required");
        config.internalDomains = [...d.internalDomains];
    }

    const messages = src.messages && typeof src.messages === "object" ? src.messages : {};
    for (const [key, limit] of Object.entries(MESSAGE_LIMITS)) {
        let text = typeof messages[key] === "string" ? messages[key].replace(/\s+/g, " ").trim() : "";
        if (!text) text = d.messages[key];
        if (text.length > limit) {
            errors.push(`Message "${key}" must be ${limit} characters or fewer`);
            text = text.slice(0, limit);
        }
        config.messages[key] = text;
    }

    return { config, errors };
}

export async function getConfig() {
    const stored = await kv.get(CONFIG_KEY);
    let parsed = null;
    try { parsed = stored ? JSON.parse(stored) : null; } catch { parsed = null; }
    if (!parsed || !parsed.config) {
        return { config: sanitizeConfig({}).config, updatedAt: null, updatedBy: null };
    }
    return { config: sanitizeConfig(parsed.config).config, updatedAt: parsed.updatedAt || null, updatedBy: parsed.updatedBy || null };
}

function changedKeys(before, after) {
    const keys = [];
    for (const key of Object.keys(after)) {
        if (key === "messages") {
            for (const m of Object.keys(after.messages)) {
                if (before.messages?.[m] !== after.messages[m]) keys.push(`messages.${m}`);
            }
        } else if (JSON.stringify(before[key]) !== JSON.stringify(after[key])) {
            keys.push(key);
        }
    }
    return keys;
}

export async function saveConfig(config, actor) {
    const current = await getConfig();
    const changes = changedKeys(current.config, config);
    const record = { config, updatedAt: new Date().toISOString(), updatedBy: actor };
    await kv.set(CONFIG_KEY, JSON.stringify(record));
    if (changes.length) {
        const entry = { at: record.updatedAt, by: actor, changes, enabled: config.enabled };
        await kv.lpush(AUDIT_KEY, JSON.stringify(entry));
        await kv.ltrim(AUDIT_KEY, 0, 499);
    }
    return { ...record, changes };
}

export async function getAudit(limit = 100) {
    const rows = await kv.lrange(AUDIT_KEY, 0, limit - 1);
    return rows.map((r) => { try { return JSON.parse(r); } catch { return null; } }).filter(Boolean);
}
