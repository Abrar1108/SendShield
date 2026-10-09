import { isConfigured, readSession } from "../../lib/auth.js";
import { hasStore } from "../../lib/store.js";
import { sendJson } from "../../lib/http.js";

export default function handler(req, res) {
    const session = readSession(req);
    sendJson(res, 200, {
        signedIn: Boolean(session),
        name: session?.name || null,
        passwordConfigured: isConfigured(),
        storage: hasStore
    }, { "Cache-Control": "no-store" });
}
