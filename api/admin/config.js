import { requireAdmin } from "../../lib/auth.js";
import { getConfig, sanitizeConfig, saveConfig } from "../../lib/config.js";
import { hasStore } from "../../lib/store.js";
import { readBody, sendJson, methodNotAllowed } from "../../lib/http.js";

export default async function handler(req, res) {
    res.setHeader("Cache-Control", "no-store");
    const session = requireAdmin(req, res);
    if (!session) return;

    if (req.method === "GET") {
        return sendJson(res, 200, { ...(await getConfig()), storage: hasStore });
    }
    if (req.method === "PUT") {
        let body;
        try { body = await readBody(req, 64 * 1024); } catch { return sendJson(res, 400, { error: "Invalid body" }); }
        const { config, errors } = sanitizeConfig(body.config);
        if (errors.length) return sendJson(res, 422, { error: "Please fix the highlighted settings.", errors });
        const saved = await saveConfig(config, session.name);
        return sendJson(res, 200, { ...saved, storage: hasStore });
    }
    methodNotAllowed(res, ["GET", "PUT"]);
}
