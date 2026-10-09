// Public, read-only settings for the add-in. Cached briefly at the edge.
import { getConfig } from "../lib/config.js";
import { allowCors, sendJson, methodNotAllowed } from "../lib/http.js";

export default async function handler(req, res) {
    allowCors(res);
    if (req.method === "OPTIONS") { res.statusCode = 204; return res.end(); }
    if (req.method !== "GET") return methodNotAllowed(res, ["GET", "OPTIONS"]);
    try {
        const { config, updatedAt } = await getConfig();
        sendJson(res, 200, { config, updatedAt }, { "Cache-Control": "public, max-age=60, s-maxage=60" });
    } catch (err) {
        console.error("config read failed", err);
        sendJson(res, 503, { error: "Settings unavailable" });
    }
}
