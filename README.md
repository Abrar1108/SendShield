# SendShield

Outlook add-in that asks users to confirm external recipients before an email leaves Toluna, with an admin dashboard for settings and reports. It replaces Standss SendGuard's external-recipient confirmation.

## How it works

1. The user clicks **Send** on an email with external recipients.
2. SendShield pauses the send and asks the user to **Review recipients**. That button opens the task pane.
3. The user ticks each external recipient (and confirms the attachments), clicks **Confirm**, then clicks **Send** again.

Design rules:

- **No loops.** Confirmations are stored per address for that draft. Adding internal or ignored recipients never re-prompts, and only a *new* external address needs confirming. After 2 blocked sends (configurable), Outlook's own **Send anyway** dialog is shown instead.
- **Fails open.** On any error, or after 4 seconds (configurable), the email is sent.
- **No network calls while sending.** Settings are fetched in the background and cached. Built-in defaults are used until the first fetch.
- **Never edits the email**, so it doesn't clash with Outreach or signature add-ins.
- **Privacy.** Reports store event counts and recipient *domains* only, never addresses, subjects or file names.

## Project layout

```
src/addin/      add-in: send handler, task pane, icons
src/admin/      dashboard (Overview, Settings, Audit log, Install)
src/manifest.xml  manifest template (URLs filled in at build time)
api/            Vercel functions: /api/config, /api/event, /api/admin/*
lib/            shared server code (settings, storage, auth, stats)
scripts/        build.mjs (Vercel build), dev.mjs (local preview)
test/           unit tests (npm test)
```

## Deploy on Vercel

1. **Import the repo:** Vercel, **Add New**, **Project**, select `SendShield`. No framework preset is needed; `vercel.json` sets the build.
2. **Environment variables** (Project, **Settings**, **Environment Variables**):
   | Name | Value |
   |---|---|
   | `ADMIN_PASSWORD` | Shared admin password, at least 12 characters |
   | `ADMIN_SESSION_SECRET` | Optional: a long random string. Changing it signs everyone out. |
3. **Storage:** Project, **Storage**, **Create Database**, **Upstash for Redis** (free plan), then connect it to the project. This adds `KV_REST_API_URL` and `KV_REST_API_TOKEN`.
4. **Redeploy** so the new variables are picked up.
5. Open `https://<your-project>.vercel.app/admin/`, sign in, and download the manifest from the **Install** tab.

The manifest and add-in files use the project's **production** URL automatically. Preview deployments sit behind Vercel's login, so test the add-in only against the production URL.

## Test in Outlook

1. Go to [aka.ms/olksideload](https://aka.ms/olksideload), then **Custom Add-ins**, **Add a custom add-in**, **Add from file**, and select the manifest.
2. Write an email to an external address and click **Send**.

For a pilot group, upload the same manifest in the Microsoft 365 admin center under **Settings**, **Integrated apps**, **Upload custom apps**.

## Local development

```bash
npm test
ADMIN_PASSWORD=local-dev-password-123 npm run dev
```

The dashboard is then at http://localhost:3000/admin/. Without Redis credentials, settings are kept in memory only.

## Notes

- Vercel's free Hobby plan is for non-commercial use; the Pro plan is needed for ongoing company use.
- The admin password sign-in is interim. Replace it with Microsoft Entra ID sign-in before production.
