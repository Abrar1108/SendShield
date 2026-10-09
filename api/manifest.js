// Serves the add-in manifest (rewritten from /manifest.xml) with the ribbon text from settings.
import { getConfig } from "../lib/config.js";
import { renderManifest } from "../lib/manifest.js";
import { resolveBaseUrl } from "../lib/baseurl.js";

export default async function handler(req, res) {
    let messages, revision = 0;
    try {
        const current = await getConfig();
        messages = current.config.messages;
        revision = current.manifestRevision;
    } catch (err) {
        console.error("manifest: settings unavailable, using defaults", err);
        const { DEFAULT_CONFIG } = await import("../lib/defaults.js");
        messages = DEFAULT_CONFIG.messages;
    }
    const env = process.env.VERCEL_ENV;
    const displayName = process.env.SENDSHIELD_DISPLAY_NAME || (env && env !== "production" ? "SendShield (preview)" : "SendShield");
    const body = renderManifest({ baseUrl: resolveBaseUrl(process.env, req.headers["x-forwarded-host"] || req.headers.host), messages, revision, displayName });
    res.statusCode = 200;
    res.setHeader("Content-Type", "application/xml; charset=utf-8");
    res.setHeader("Cache-Control", "no-store");
    if (new URL(req.url, "http://x").searchParams.has("download")) {
        res.setHeader("Content-Disposition", 'attachment; filename="sendshield-manifest.xml"');
    }
    res.end(body);
}
