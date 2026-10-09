// Built-in settings. The add-in embeds these at build time and uses them until it has fetched
// settings from /api/config, so SendShield works even if the server or storage is unavailable.
export const DEFAULT_CONFIG = {
    enabled: true,
    internalDomains: ["toluna.com", "metrixlab.com"],
    includeSubdomains: true,
    trustedDomains: [],
    ignoredAddresses: [],
    confirmAttachments: true,
    showRecipientBanner: true,
    failOpenTimeoutMs: 4000,
    maxBlocksBeforeFallback: 2,
    telemetryEnabled: true,
    messages: {
        blockPrompt: "Please confirm {count} external recipient(s): {list}. {hint}",
        attachmentPrompt: "Please confirm the {attachments} attachment(s) going outside Toluna. {hint}",
        fallbackPrompt: "You are sending outside Toluna to: {list}. Choose Send Anyway only if every recipient is correct.",
        reviewButton: "Review recipients",
        banner: "{count} external recipient(s): {list}",
        taskpaneIntro: "This email is going outside Toluna. Tick each external recipient to confirm it is correct."
    }
};

export const MESSAGE_LIMITS = {
    blockPrompt: 200,
    attachmentPrompt: 200,
    fallbackPrompt: 200,
    reviewButton: 20,
    banner: 120,
    taskpaneIntro: 300
};
