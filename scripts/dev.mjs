// Local preview that mimics Vercel: serves public/ and routes /api/* to the functions in api/.
// Usage: ADMIN_PASSWORD=some-long-password npm run dev   (then open http://localhost:3000/admin/)
import { createServer } from "node:http";
import { existsSync, readFileSync, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { execFileSync } from "node:child_process";

const root = join(fileURLToPath(import.meta.url), "..", "..");
const port = Number(process.env.PORT || 3000);
execFileSync(process.execPath, [join(root, "scripts", "build.mjs")], { stdio: "inherit", env: { ...process.env, SENDSHIELD_BASE_URL: `http://localhost:${port}` } });

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".xml": "application/xml; charset=utf-8", ".png": "image/png", ".json": "application/json" };
const vercel = JSON.parse(readFileSync(join(root, "vercel.json"), "utf8"));

function applyHeaders(path, res) {
    for (const rule of vercel.headers) {
        if (new RegExp("^" + rule.source.replace("(.*)", ".*") + "$").test(path)) {
            for (const h of rule.headers) res.setHeader(h.key, h.value);
        }
    }
}

createServer(async (req, res) => {
    const url = new URL(req.url, `http://localhost:${port}`);
    let path = decodeURIComponent(url.pathname);
    for (const rule of vercel.rewrites || []) if (rule.source === path) path = rule.destination;
    try {
        if (path.startsWith("/api/")) {
            const file = join(root, normalize(path).replace(/\/$/, "") + ".js");
            if (!file.startsWith(join(root, "api")) || !existsSync(file)) { res.statusCode = 404; return res.end("Not found"); }
            const mod = await import(pathToFileURL(file).href);
            return await mod.default(req, res);
        }
        if (path === "/admin") { res.writeHead(307, { Location: "/admin/" }); return res.end(); }
        if (path.endsWith("/")) path += "index.html";
        const file = join(root, "public", normalize(path));
        if (!file.startsWith(join(root, "public")) || !existsSync(file) || statSync(file).isDirectory()) { res.statusCode = 404; return res.end("Not found"); }
        applyHeaders(path, res);
        res.setHeader("Content-Type", TYPES[extname(file)] || "application/octet-stream");
        res.end(readFileSync(file));
    } catch (err) {
        console.error(err);
        if (!res.headersSent) res.statusCode = 500;
        res.end("Server error");
    }
}).listen(port, () => console.log(`SendShield dev server on http://localhost:${port}/admin/`));
