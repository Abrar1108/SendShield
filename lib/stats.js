// Usage counters. Only event types, counts and recipient DOMAINS are stored - never addresses,
// subjects or attachment names.
import { kv, hasStore } from "./store.js";

export const EVENT_TYPES = ["prompt_shown", "taskpane_confirmed", "sent_confirmed", "loop_breaker", "fail_open"];
const RETENTION_SECONDS = 400 * 24 * 3600;
const DOMAIN_RE = /^(?=.{3,253}$)([a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/;

export function dayKey(date = new Date()) {
    return date.toISOString().slice(0, 10);
}

export function validateEvent(body) {
    if (!body || typeof body !== "object") return null;
    if (!EVENT_TYPES.includes(body.type)) return null;
    const domains = Array.isArray(body.externalDomains)
        ? [...new Set(body.externalDomains.map((d) => String(d).toLowerCase().trim()).filter((d) => DOMAIN_RE.test(d)))].slice(0, 20)
        : [];
    const platform = String(body.platform || "").replace(/[^A-Za-z0-9_-]/g, "").slice(0, 20) || "unknown";
    return { type: body.type, domains, platform };
}

export async function recordEvent(event) {
    if (!hasStore) return;
    const day = dayKey();
    const statsKey = `ss:stats:${day}`;
    const commands = [
        ["HINCRBY", statsKey, event.type, 1],
        ["HINCRBY", statsKey, `platform:${event.platform}`, 1],
        ["EXPIRE", statsKey, RETENTION_SECONDS]
    ];
    // Domain counts only for emails that actually went out after confirmation.
    if (event.type === "sent_confirmed" && event.domains.length) {
        const domainKey = `ss:domains:${day}`;
        for (const d of event.domains) commands.push(["ZINCRBY", domainKey, 1, d]);
        commands.push(["EXPIRE", domainKey, RETENTION_SECONDS]);
    }
    await kv.pipeline(commands);
}

function pairs(flat) {
    const out = [];
    for (let i = 0; i + 1 < (flat || []).length; i += 2) out.push([flat[i], Number(flat[i + 1]) || 0]);
    return out;
}

export async function getReport(days) {
    const span = Math.min(90, Math.max(1, days | 0));
    const dates = [];
    const today = new Date();
    for (let i = span - 1; i >= 0; i--) {
        const d = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate() - i));
        dates.push(dayKey(d));
    }
    if (!hasStore) return { storage: false, days: dates.map((date) => ({ date })), totals: {}, platforms: {}, topDomains: [] };

    const commands = [];
    for (const date of dates) {
        commands.push(["HGETALL", `ss:stats:${date}`]);
        commands.push(["ZREVRANGE", `ss:domains:${date}`, 0, 99, "WITHSCORES"]);
    }
    const results = await kv.pipeline(commands);

    const totals = Object.fromEntries(EVENT_TYPES.map((t) => [t, 0]));
    const platforms = {};
    const domainTotals = new Map();
    const daily = dates.map((date, i) => {
        const row = { date };
        for (const t of EVENT_TYPES) row[t] = 0;
        for (const [field, value] of pairs(results[i * 2])) {
            if (field.startsWith("platform:")) {
                const p = field.slice(9);
                platforms[p] = (platforms[p] || 0) + value;
            } else if (field in totals) {
                row[field] = value;
                totals[field] += value;
            }
        }
        for (const [domain, value] of pairs(results[i * 2 + 1])) {
            domainTotals.set(domain, (domainTotals.get(domain) || 0) + value);
        }
        return row;
    });

    const topDomains = [...domainTotals.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 25)
        .map(([domain, count]) => ({ domain, count }));

    return { storage: true, days: daily, totals, platforms, topDomains };
}
