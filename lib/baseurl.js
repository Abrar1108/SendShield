// The public URL of this deployment. Used for the add-in files (build time) and the manifest (runtime).
export function resolveBaseUrl(env = process.env, host) {
    let base = env.SENDSHIELD_BASE_URL;
    if (!base && env.VERCEL_ENV === "production" && env.VERCEL_PROJECT_PRODUCTION_URL) base = env.VERCEL_PROJECT_PRODUCTION_URL;
    if (!base && env.VERCEL_URL) base = env.VERCEL_URL;
    if (!base && host) base = host;
    if (!base) base = "http://localhost:3000";
    if (!/^https?:\/\//.test(base)) base = (base.startsWith("localhost") ? "http://" : "https://") + base;
    return base.replace(/\/+$/, "");
}
