// Built-in settings. The add-in embeds these at build time and uses them until it has fetched
// settings from /api/config, so SendShield works even if the server or storage is unavailable.

// Every piece of user-facing text. The dashboard renders its message editor from this list.
// `where` groups fields in the editor; `vars` lists the placeholders each message supports.
export const MESSAGE_FIELDS = [
    { key: "blockPrompt", where: "dialog", label: "Send prompt", limit: 200, multiline: true, vars: ["count", "list", "hint"],
        value: "Please confirm {count} external recipient(s): {list}. {hint}" },
    { key: "attachmentPrompt", where: "dialog", label: "Attachment prompt", limit: 200, multiline: true, vars: ["attachments", "hint"],
        value: "Please confirm the {attachments} attachment(s) going outside Toluna. {hint}" },
    { key: "hintModern", where: "dialog", label: "Instruction (current Outlook)", limit: 120, multiline: true, vars: ["button"],
        value: "Click \"{button}\" to review them, then click Send again." },
    { key: "hintLegacy", where: "dialog", label: "Instruction (older Outlook)", limit: 120, multiline: true, vars: ["ribbon"],
        value: "Click {ribbon} on the ribbon, then click Send again." },
    { key: "reviewButton", where: "dialog", label: "Review button", limit: 20, vars: [],
        value: "Review recipients" },
    { key: "fallbackPrompt", where: "fallback", label: "Safety-net prompt", limit: 200, multiline: true, vars: ["count", "list"],
        value: "You are sending outside Toluna to: {list}. Choose Send Anyway only if every recipient is correct." },
    { key: "banner", where: "banner", label: "Compose banner", limit: 120, vars: ["count", "list"],
        value: "{count} external recipient(s): {list}" },
    { key: "taskpaneTitle", where: "taskpane", label: "Title", limit: 40, vars: [],
        value: "SendShield" },
    { key: "taskpaneIntro", where: "taskpane", label: "Introduction", limit: 300, multiline: true, vars: [],
        value: "This email is going outside Toluna. Tick each external recipient to confirm it is correct." },
    { key: "recipientsHeading", where: "taskpane", label: "Recipients heading", limit: 40, vars: [],
        value: "External recipients" },
    { key: "selectAll", where: "taskpane", label: "Select all", limit: 30, vars: [],
        value: "Select all" },
    { key: "confirmAllCheckbox", where: "taskpane", label: "Confirm-all checkbox (single tick style)", limit: 150, multiline: true, vars: ["count"],
        value: "I confirm all {count} external recipient(s) are correct." },
    { key: "attachmentsHeading", where: "taskpane", label: "Attachments heading", limit: 40, vars: [],
        value: "Attachments" },
    { key: "attachmentsCheckbox", where: "taskpane", label: "Attachments checkbox", limit: 200, multiline: true, vars: [],
        value: "I have checked these attachments are correct to send outside Toluna." },
    { key: "confirmButton", where: "taskpane", label: "Confirm button", limit: 30, vars: [],
        value: "Confirm" },
    { key: "remainingRecipients", where: "taskpane", label: "Recipients left to tick", limit: 80, vars: ["count"],
        value: "{count} recipient(s) still to tick" },
    { key: "remainingAttachments", where: "taskpane", label: "Attachments not confirmed", limit: 80, vars: [],
        value: "Please confirm the attachments" },
    { key: "confirmedMessage", where: "taskpane", label: "Confirmed message", limit: 150, multiline: true, vars: [],
        value: "Confirmed. You can now click Send." },
    { key: "saveFailed", where: "taskpane", label: "Save failed", limit: 200, multiline: true, vars: [],
        value: "Couldn't save your confirmation. You can still send - Outlook will ask you to confirm in a dialog." },
    { key: "noExternal", where: "taskpane", label: "No external recipients", limit: 150, multiline: true, vars: [],
        value: "No external recipients. Nothing to confirm." },
    { key: "legacyNotice", where: "taskpane", label: "Older Outlook notice", limit: 200, multiline: true, vars: [],
        value: "Your Outlook version can't store confirmations. When you click Send you'll be asked to confirm in a dialog instead." },
    { key: "pilotInactive", where: "taskpane", label: "Not in pilot", limit: 200, multiline: true, vars: [],
        value: "SendShield is being piloted and is not active for your mailbox yet." },
    { key: "internalSummary", where: "taskpane", label: "Internal recipients footer", limit: 60, vars: ["count"],
        value: "{count} internal recipient(s)" },
    { key: "ribbonGroup", where: "ribbon", label: "Ribbon group", limit: 32, vars: [], manifest: true,
        value: "SendShield" },
    { key: "ribbonButton", where: "ribbon", label: "Ribbon button", limit: 32, vars: [], manifest: true,
        value: "Confirm recipients" },
    { key: "ribbonTooltip", where: "ribbon", label: "Ribbon tooltip", limit: 250, multiline: true, vars: [], manifest: true,
        value: "Review and confirm external recipients before sending." }
];

export const MESSAGE_LIMITS = Object.fromEntries(MESSAGE_FIELDS.map((f) => [f.key, f.limit]));
export const MANIFEST_MESSAGE_KEYS = MESSAGE_FIELDS.filter((f) => f.manifest).map((f) => f.key);

export const DEFAULT_CONFIG = {
    enabled: true,
    pilotMode: false,
    pilotUsers: [],
    internalDomains: ["toluna.com", "metrixlab.com"],
    includeSubdomains: true,
    trustedDomains: [],
    ignoredAddresses: [],
    confirmStyle: "each",          // "each" = tick every recipient, "all" = one confirm-all tick
    minExternalRecipients: 1,
    confirmAttachments: true,
    showRecipientBanner: true,
    failOpenTimeoutMs: 4000,
    maxBlocksBeforeFallback: 2,
    configRefreshMinutes: 15,
    telemetryEnabled: true,
    messages: Object.fromEntries(MESSAGE_FIELDS.map((f) => [f.key, f.value]))
};
