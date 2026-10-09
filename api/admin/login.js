import { checkPassword, isConfigured, sessionCookie } from "../../lib/auth.js";
import { clientIp, readBody, sendJson, methodNotAllowed } from "../../lib/http.js";
import { kv } from "../../lib/store.js";

const MAX_ATTEMPTS = 10;
const WINDOW_SECONDS = 15 * 60;

export default async function handler(req, res) {
    if (req.method !== "POST") return methodNotAllowed(res, ["POST"]);
    if (req.headers["x-sendshield"] !== "1") return sendJson(res, 403, { error: "Missing request header" });
    if (!isConfigured()) {
        return sendJson(res, 503, { error: "ADMIN_PASSWORD is not set (minimum 12 characters). Add it in Vercel > Settings > Environment Variables and redeploy." });
    }

    const attempts = await kv.incrWithExpiry(`ss:login:${clientIp(req)}`, WINDOW_SECONDS);
    if (attempts > MAX_ATTEMPTS) return sendJson(res, 429, { error: "Too many attempts. Try again in 15 minutes." });

    let body;
    try { body = await readBody(req, 2048); } catch { return sendJson(res, 400, { error: "Invalid body" }); }
    const name = String(body.name || "").replace(/[^\p{L}\p{N} .'@_-]/gu, "").trim().slice(0, 60);
    if (!name) return sendJson(res, 400, { error: "Please enter your name (used in the audit log)." });
    if (!checkPassword(String(body.password || ""))) return sendJson(res, 401, { error: "Incorrect password." });

    sendJson(res, 200, { name }, { "Set-Cookie": sessionCookie(name) });
}
