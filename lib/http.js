export function sendJson(res, status, body, headers = {}) {
    res.statusCode = status;
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    for (const [k, v] of Object.entries(headers)) res.setHeader(k, v);
    res.end(JSON.stringify(body));
}

// Public endpoints are called from inside Outlook (different origin), so they allow any origin.
// They never use cookies, so this does not expose admin endpoints.
export function allowCors(res) {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    res.setHeader("Access-Control-Max-Age", "86400");
}

export async function readBody(req, maxBytes = 16 * 1024) {
    if (req.body !== undefined && req.body !== null) {
        if (typeof req.body === "string") return parse(req.body, maxBytes);
        if (Buffer.isBuffer(req.body)) return parse(req.body.toString("utf8"), maxBytes);
        return req.body;
    }
    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
        size += chunk.length;
        if (size > maxBytes) throw new Error("Body too large");
        chunks.push(chunk);
    }
    return parse(Buffer.concat(chunks).toString("utf8"), maxBytes);
}

function parse(text, maxBytes) {
    if (text.length > maxBytes) throw new Error("Body too large");
    if (!text) return {};
    return JSON.parse(text);
}

export function clientIp(req) {
    const fwd = String(req.headers["x-forwarded-for"] || "");
    return fwd.split(",")[0].trim() || req.socket?.remoteAddress || "unknown";
}

export function methodNotAllowed(res, allowed) {
    res.setHeader("Allow", allowed.join(", "));
    sendJson(res, 405, { error: "Method not allowed" });
}
