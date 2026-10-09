import { requireAdmin } from "../../lib/auth.js";
import { getAudit } from "../../lib/config.js";
import { sendJson, methodNotAllowed } from "../../lib/http.js";

export default async function handler(req, res) {
    res.setHeader("Cache-Control", "no-store");
    if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
    if (!requireAdmin(req, res)) return;
    sendJson(res, 200, { entries: await getAudit(200) });
}
