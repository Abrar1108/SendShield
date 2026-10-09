// SendShield event handlers (OnMessageSend, OnMessageRecipientsChanged).
// The build concatenates core.js + runtime.js + this file into addin/handler.js, because the classic
// Outlook JS-only runtime loads exactly one script file.
(function (root) {
    "use strict";

    var Core = root.SendShieldCore;
    var RT = root.SendShieldRuntime;
    var DEFAULT_TIMEOUT_MS = 4000;
    var MAX_PROMPT = 250;  // Outlook limit for Smart Alerts dialog text (kept conservative).
    var MAX_BANNER = 150;  // Outlook limit for notification messages.
    var REVIEW_COMMAND_ID = "msgComposeOpenPaneButton";

    // Warm the settings cache as soon as the runtime loads.
    try { RT.refreshConfig(false); } catch (e) { /* ignore */ }

    function details(decision) {
        var cls = decision.classification || { externals: [], externalDomains: [] };
        return {
            externalCount: cls.externals.length,
            externalDomains: cls.externalDomains,
            attachmentCount: (decision.attachments || []).length,
            reason: decision.reason
        };
    }

    function onMessageSendHandler(event) {
        var completed = false;
        var config = null;
        var item = Office.context.mailbox.item;

        function finish(options, telemetryType, decision) {
            if (completed) return;
            completed = true;
            clearTimeout(timer);
            if (telemetryType) {
                try { RT.sendEvent(config, telemetryType, decision ? details(decision) : { reason: "timeout" }); } catch (e) { /* ignore */ }
            }
            event.completed(options);
        }

        // Fail open: if anything is slow or breaks, the email is sent.
        var timer = setTimeout(function () {
            finish({ allowEvent: true }, "fail_open", null);
        }, DEFAULT_TIMEOUT_MS);

        try {
            RT.loadConfig(function (cfg) {
                config = cfg;
                if (completed) return;
                if (!cfg.enabled) { finish({ allowEvent: true }); return; }

                // Use the fail-open window from settings (validated server-side to 1-8 seconds).
                if (cfg.failOpenTimeoutMs && cfg.failOpenTimeoutMs !== DEFAULT_TIMEOUT_MS) {
                    clearTimeout(timer);
                    timer = setTimeout(function () { finish({ allowEvent: true }, "fail_open", null); }, cfg.failOpenTimeoutMs);
                }

                var recipients = null, recipientsOk = true, attachments = [], session = {}, sessionOk = false;
                var remaining = 3;
                function step() {
                    if (--remaining > 0 || completed) return;
                    if (!recipientsOk) { finish({ allowEvent: true }, "fail_open", { reason: "recipients_unreadable" }); return; }
                    decide(cfg, recipients, attachments, session, sessionOk);
                }

                RT.getRecipients(item, function (list, ok) { recipients = list; recipientsOk = ok; step(); });
                if (cfg.confirmAttachments) RT.getAttachmentNames(item, function (names) { attachments = names; step(); });
                else step();
                RT.getSession(item, function (values, ok) { session = values; sessionOk = ok; step(); });
            });
        } catch (e) {
            finish({ allowEvent: true }, "fail_open", { reason: "exception" });
        }

        function decide(cfg, recipients, attachments, session, sessionOk) {
            var decision = Core.evaluateSend({
                config: cfg,
                recipients: recipients,
                attachments: attachments,
                session: session,
                sessionAvailable: sessionOk
            });

            if (decision.action === "allow") {
                finish({ allowEvent: true }, decision.reason === "confirmed" ? "sent_confirmed" : null, decision);
                return;
            }

            var externals = decision.classification.externals;
            var modern = RT.supports("1.14");

            if (decision.action === "fallback") {
                if (!modern) {
                    // Older Outlook can't show "Send anyway" from a soft block, so fail open.
                    finish({ allowEvent: true }, "loop_breaker", decision);
                    return;
                }
                var fallbackText = Core.fillTemplate(cfg.messages.fallbackPrompt, {
                    count: externals.length,
                    list: Core.formatList(externals, 3)
                });
                var overrideMode = (Office.MailboxEnums && Office.MailboxEnums.SendModeOverride && Office.MailboxEnums.SendModeOverride.PromptUser) || "promptUser";
                finish({
                    allowEvent: false,
                    errorMessage: Core.truncate(fallbackText, MAX_PROMPT),
                    sendModeOverride: overrideMode
                }, "loop_breaker", decision);
                return;
            }

            // action === "block": ask the user to confirm in the task pane.
            var pendingR = decision.pendingRecipients;
            var hint = modern
                ? "Click \"" + cfg.messages.reviewButton + "\" to review them, then click Send again."
                : "Open SendShield > Confirm recipients on the ribbon, then click Send again.";
            var text = Core.fillTemplate(cfg.messages.blockPrompt, {
                count: pendingR.length || externals.length,
                list: Core.formatList(pendingR.length ? pendingR : externals, 3),
                attachments: decision.pendingAttachments.length,
                hint: hint
            });
            if (!pendingR.length && decision.pendingAttachments.length) {
                text = Core.fillTemplate(cfg.messages.attachmentPrompt, {
                    attachments: decision.pendingAttachments.length,
                    hint: hint
                });
            }

            var options = { allowEvent: false, errorMessage: Core.truncate(text, MAX_PROMPT) };
            if (modern) {
                options.cancelLabel = Core.truncate(cfg.messages.reviewButton, 20);
                options.commandId = REVIEW_COMMAND_ID;
            }

            RT.setSession(item, { ssBlocks: decision.blocks }, function () {
                finish(options, "prompt_shown", decision);
            });
        }
    }

    function onRecipientsChangedHandler(event) {
        var finished = false;
        function done() { if (!finished) { finished = true; event.completed(); } }
        var guard = setTimeout(done, 4000);

        try {
            RT.refreshConfig(false);
            var item = Office.context.mailbox.item;
            RT.loadConfig(function (cfg) {
                if (!cfg.enabled || !cfg.showRecipientBanner) {
                    item.notificationMessages.removeAsync("ssExternal", function () { clearTimeout(guard); done(); });
                    return;
                }
                RT.getRecipients(item, function (list) {
                    var cls = Core.classifyRecipients(list, cfg);
                    if (cls.externals.length === 0) {
                        item.notificationMessages.removeAsync("ssExternal", function () { clearTimeout(guard); done(); });
                        return;
                    }
                    var message = Core.truncate(Core.fillTemplate(cfg.messages.banner, {
                        count: cls.externals.length,
                        list: Core.formatList(cls.externals, 2)
                    }), MAX_BANNER);
                    item.notificationMessages.replaceAsync("ssExternal", {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: message,
                        icon: "Icon.16x16",
                        persistent: false
                    }, function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) { clearTimeout(guard); done(); return; }
                        // Some clients reject informational banners from event runtimes; fall back to the plain style.
                        item.notificationMessages.replaceAsync("ssExternal", {
                            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                            message: message
                        }, function () { clearTimeout(guard); done(); });
                    });
                });
            });
        } catch (e) {
            clearTimeout(guard);
            done();
        }
    }

    Office.onReady(function () {});

    try {
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
        Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
    } catch (e) { /* older hosts use the globals below */ }

    root.onMessageSendHandler = onMessageSendHandler;
    root.onRecipientsChangedHandler = onRecipientsChangedHandler;
})(typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : this);
