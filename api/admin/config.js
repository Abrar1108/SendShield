import { requireAdmin } from "../../lib/auth.js";
import { getConfig, sanitizeConfig, saveConfig } from "../../lib/config.js";
import { DEFAULT_CONFIG, MESSAGE_FIELDS } from "../../lib/defaults.js";
import { manifestVersion } from "../../lib/manifest.js";
import { hasStore } from "../../lib/store.js";
import { readBody, sendJson, methodNotAllowed } from "../../lib/http.js";

// Everything the dashboard needs to render the settings and message editors.
function describe(record) {
    return {
        config: record.config,
        updatedAt: record.updatedAt,
        updatedBy: record.updatedBy,
        manifestVersion: manifestVersion(record.manifestRevision),
        manifestChanged: Boolean(record.manifestChanged),
        defaults: DEFAULT_CONFIG,
        messageFields: MESSAGE_FIELDS.map(({ value, ...field }) => field),
        storage: hasStore
    };
}

export default async function handler(req, res) {
    res.setHeader("Cache-Control", "no-store");
    const session = requireAdmin(req, res);
    if (!session) return;

    if (req.method === "GET") {
        return sendJson(res, 200, describe(await getConfig()));
    }
    if (req.method === "PUT") {
        let body;
        try { body = await readBody(req, 64 * 1024); } catch { return sendJson(res, 400, { error: "Invalid body" }); }
        const { config, errors } = sanitizeConfig(body.config);
        if (errors.length) return sendJson(res, 422, { error: "Please fix the settings below.", errors });
        return sendJson(res, 200, describe(await saveConfig(config, session.name)));
    }
    methodNotAllowed(res, ["GET", "PUT"]);
}
