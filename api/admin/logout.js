import { clearCookie } from "../../lib/auth.js";
import { sendJson, methodNotAllowed } from "../../lib/http.js";

export default function handler(req, res) {
    if (req.method !== "POST") return methodNotAllowed(res, ["POST"]);
    sendJson(res, 200, { ok: true }, { "Set-Cookie": clearCookie() });
}
