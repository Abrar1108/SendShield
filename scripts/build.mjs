// Builds the static site into public/:
//  - copies src/ to public/
//  - fills placeholders (__BASE_URL__, __VERSION__, __DEFAULT_CONFIG__ ...) in text files
//  - bundles core.js + runtime.js + handler.js into addin/handler.js (the classic Outlook
//    JS-only runtime loads a single file)
//
// The base URL comes from SENDSHIELD_BASE_URL, otherwise from Vercel's system variables.
import { cpSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, extname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { DEFAULT_CONFIG } from "../lib/defaults.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const src = join(root, "src");
const out = join(root, "public");
const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));

function resolveBaseUrl() {
    const env = process.env;
    let base = env.SENDSHIELD_BASE_URL;
    if (!base && env.VERCEL_ENV === "production" && env.VERCEL_PROJECT_PRODUCTION_URL) base = env.VERCEL_PROJECT_PRODUCTION_URL;
    if (!base && env.VERCEL_URL) base = env.VERCEL_URL;
    if (!base) base = "http://localhost:3000";
    if (!/^https?:\/\//.test(base)) base = "https://" + base;
    return base.replace(/\/+$/, "");
}

const baseUrl = resolveBaseUrl();
const version = pkg.version;
const manifestVersion = version.split(".").concat(["0", "0", "0", "0"]).slice(0, 4).join(".");
const isProduction = process.env.VERCEL_ENV === "production";

const replacements = {
    __BASE_URL__: baseUrl,
    __VERSION__: version,
    __MANIFEST_VERSION__: manifestVersion,
    __ADDIN_ID__: process.env.SENDSHIELD_ADDIN_ID || pkg.sendshield.addinId,
    __DISPLAY_NAME__: process.env.SENDSHIELD_DISPLAY_NAME || (isProduction || !process.env.VERCEL_ENV ? "SendShield" : "SendShield (preview)"),
    __DEFAULT_CONFIG__: JSON.stringify(DEFAULT_CONFIG),
    __BUILD_TIME__: new Date().toISOString()
};

const TEXT = new Set([".html", ".js", ".css", ".xml", ".json", ".txt"]);

function fill(text) {
    let result = text;
    for (const [key, value] of Object.entries(replacements)) result = result.split(key).join(value);
    return result;
}

function walk(dir) {
    const files = [];
    for (const name of readdirSync(dir)) {
        const full = join(dir, name);
        if (statSync(full).isDirectory()) files.push(...walk(full));
        else files.push(full);
    }
    return files;
}

rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
cpSync(src, out, { recursive: true });

for (const file of walk(out)) {
    if (TEXT.has(extname(file))) writeFileSync(file, fill(readFileSync(file, "utf8")));
}

// Single-file bundle for the event runtime; the task pane still loads the individual files.
const bundle = ["core.js", "runtime.js", "handler.js"]
    .map((name) => `// ---- ${name}\n` + readFileSync(join(out, "addin", name), "utf8"))
    .join("\n");
writeFileSync(join(out, "addin", "handler.js"), bundle);

const leftovers = walk(out).filter((f) => TEXT.has(extname(f)) && /__[A-Z_]+__/.test(readFileSync(f, "utf8")));
if (leftovers.length) {
    console.error("Unfilled placeholders in:", leftovers);
    process.exit(1);
}

console.log(`SendShield ${version} built for ${baseUrl} (manifest ${manifestVersion})`);
