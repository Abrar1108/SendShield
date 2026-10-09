// Tiny Redis client for Upstash's REST API (what Vercel's Marketplace "Upstash for Redis" provides).
// No npm dependency: commands are sent as JSON arrays over HTTPS.
// Without credentials it falls back to in-memory storage so the app still runs (nothing persists).

const URL_ = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL || "";
const TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN || "";

export const hasStore = Boolean(URL_ && TOKEN);

async function call(path, body) {
    const response = await fetch(URL_.replace(/\/$/, "") + path, {
        method: "POST",
        headers: { Authorization: `Bearer ${TOKEN}`, "Content-Type": "application/json" },
        body: JSON.stringify(body)
    });
    const json = await response.json();
    if (!response.ok || json.error) throw new Error(`Redis error: ${json.error || response.status}`);
    return json;
}

async function command(...args) {
    return (await call("", args)).result;
}

async function pipeline(commands) {
    if (!commands.length) return [];
    const results = await call("/pipeline", commands);
    return results.map((r) => (r.error ? null : r.result));
}

// ------------------------------------------------------------------ in-memory fallback
const mem = globalThis.__sendshieldMem || (globalThis.__sendshieldMem = { strings: new Map(), lists: new Map(), counters: new Map() });

const memoryKv = {
    async get(key) { return mem.strings.get(key) ?? null; },
    async set(key, value) { mem.strings.set(key, value); return "OK"; },
    async lpush(key, value) { const l = mem.lists.get(key) || []; l.unshift(value); mem.lists.set(key, l); return l.length; },
    async ltrim(key, start, stop) { const l = mem.lists.get(key) || []; mem.lists.set(key, l.slice(start, stop + 1)); return "OK"; },
    async lrange(key, start, stop) { return (mem.lists.get(key) || []).slice(start, stop + 1); },
    async incrWithExpiry(key) { const n = (mem.counters.get(key) || 0) + 1; mem.counters.set(key, n); return n; },
    async pipeline() { return []; }
};

const redisKv = {
    get: (key) => command("GET", key),
    set: (key, value) => command("SET", key, value),
    lpush: (key, value) => command("LPUSH", key, value),
    ltrim: (key, start, stop) => command("LTRIM", key, start, stop),
    lrange: (key, start, stop) => command("LRANGE", key, start, stop),
    async incrWithExpiry(key, seconds) {
        const [count] = await pipeline([["INCR", key], ["EXPIRE", key, seconds]]);
        return count;
    },
    pipeline
};

export const kv = hasStore ? redisKv : memoryKv;
