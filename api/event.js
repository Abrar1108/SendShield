// Telemetry from the add-in: event type, counts and recipient domains only.
import { recordEvent, validateEvent } from "../lib/stats.js";
import { allowCors, readBody, sendJson, methodNotAllowed } from "../lib/http.js";

export default async function handler(req, res) {
    allowCors(res);
    if (req.method === "OPTIONS") { res.statusCode = 204; return res.end(); }
    if (req.method !== "POST") return methodNotAllowed(res, ["POST", "OPTIONS"]);
    let event;
    try {
        event = validateEvent(await readBody(req, 4096));
    } catch {
        return sendJson(res, 400, { error: "Invalid body" });
    }
    if (!event) return sendJson(res, 400, { error: "Invalid event" });
    try {
        await recordEvent(event);
    } catch (err) {
        console.error("event write failed", err);
    }
    res.statusCode = 204;
    res.end();
}
