import { requireAdmin } from "../../lib/auth.js";
import { getReport } from "../../lib/stats.js";
import { sendJson, methodNotAllowed } from "../../lib/http.js";

export default async function handler(req, res) {
    res.setHeader("Cache-Control", "no-store");
    if (req.method !== "GET") return methodNotAllowed(res, ["GET"]);
    if (!requireAdmin(req, res)) return;
    const days = Number.parseInt(new URL(req.url, "http://x").searchParams.get("days") || "30", 10);
    try {
        sendJson(res, 200, await getReport(days));
    } catch (err) {
        console.error("report failed", err);
        sendJson(res, 503, { error: "Reports unavailable" });
    }
}
