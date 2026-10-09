// Interim admin login for testing: one shared password (ADMIN_PASSWORD) + the admin's name for the audit log.
// Sessions are HMAC-signed cookies. Replace with Microsoft Entra ID sign-in for production.
import crypto from "node:crypto";
import { sendJson } from "./http.js";

const COOKIE = "ss_admin";
const SESSION_HOURS = 8;

function sha256(text) {
    return crypto.createHash("sha256").update(String(text)).digest();
}

function secret() {
    return process.env.ADMIN_SESSION_SECRET || sha256("sendshield-session:" + (process.env.ADMIN_PASSWORD || "")).toString("hex");
}

function sign(payload) {
    return crypto.createHmac("sha256", secret()).update(payload).digest("base64url");
}

export function isConfigured() {
    return Boolean(process.env.ADMIN_PASSWORD && process.env.ADMIN_PASSWORD.length >= 12);
}

export function checkPassword(password) {
    if (!isConfigured()) return false;
    return crypto.timingSafeEqual(sha256(password || ""), sha256(process.env.ADMIN_PASSWORD));
}

export function sessionCookie(name) {
    const payload = Buffer.from(JSON.stringify({ n: name, exp: Date.now() + SESSION_HOURS * 3600 * 1000 })).toString("base64url");
    const token = `${payload}.${sign(payload)}`;
    return `${COOKIE}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${SESSION_HOURS * 3600}`;
}

export function clearCookie() {
    return `${COOKIE}=; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=0`;
}

export function readSession(req) {
    if (!isConfigured()) return null;
    const cookies = String(req.headers.cookie || "").split(";").map((c) => c.trim());
    const raw = cookies.find((c) => c.startsWith(COOKIE + "="));
    if (!raw) return null;
    const [payload, signature] = raw.slice(COOKIE.length + 1).split(".");
    if (!payload || !signature) return null;
    const expected = sign(payload);
    if (expected.length !== signature.length || !crypto.timingSafeEqual(Buffer.from(expected), Buffer.from(signature))) return null;
    try {
        const data = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
        if (!data.exp || data.exp < Date.now()) return null;
        return { name: data.n };
    } catch {
        return null;
    }
}

// Admin API guard. Mutating requests must also carry a custom header, which a cross-site form
// or script cannot send without a CORS pre-flight (that admin endpoints never approve).
export function requireAdmin(req, res) {
    const session = readSession(req);
    if (!session) {
        sendJson(res, 401, { error: "Not signed in" });
        return null;
    }
    if (req.method !== "GET" && req.headers["x-sendshield"] !== "1") {
        sendJson(res, 403, { error: "Missing request header" });
        return null;
    }
    return session;
}
