// SendShield event handlers (OnMessageSend, OnMessageRecipientsChanged).
// The build concatenates core.js + runtime.js + this file into addin/handler.js, because the classic
// Outlook JS-only runtime loads exactly one script file.
//
// Coexistence with other add-ins (Outreach, signature tools...): this code only READS the email.
// Its only writes are to its own per-add-in session data and its own banner, it makes no network
// call before deciding, and it always completes - so it can't block or overwrite another add-in.
(function (root) {
    "use strict";

    var Core = root.SendShieldCore;
    var RT = root.SendShieldRuntime;
    var DEFAULT_TIMEOUT_MS = 4000;
    var MAX_PROMPT = 250;  // Outlook limit for Smart Alerts dialog text (kept conservative).
    var MAX_BANNER = 150;  // Outlook limit for notification messages.
    var REVIEW_COMMAND_ID = "msgComposeOpenPaneButton";
    var BANNER_KEY = "ssExternal";

    // Warm the settings cache as soon as the runtime loads.
    try { RT.refreshConfig(false); } catch (e) { /* ignore */ }

    function details(decision) {
        var cls = (decision && decision.classification) || { externals: [], externalDomains: [] };
        return {
            externalCount: cls.externals.length,
            externalDomains: cls.externalDomains,
            attachmentCount: ((decision && decision.attachments) || []).length,
            reason: decision ? decision.reason : "timeout"
        };
    }

    function onMessageSendHandler(event) {
        var completed = false;
        var config = null;
        var item = Office.context.mailbox.item;
        var userEmail = RT.userEmail();

        function finish(options, telemetryType, decision) {
            if (completed) return;
            completed = true;
            clearTimeout(timer);
            if (telemetryType) {
                try { RT.sendEvent(config, telemetryType, details(decision)); } catch (e) { /* ignore */ }
            }
            event.completed(options);
        }

        // Fail open: if anything is slow or breaks, the email is sent.
        var timer = setTimeout(function () { finish({ allowEvent: true }, "fail_open", null); }, DEFAULT_TIMEOUT_MS);

        function evaluate(recipients, attachments, session, sessionOk) {
            return Core.evaluateSend({
                config: config,
                userEmail: userEmail,
                recipients: recipients,
                attachments: attachments,
                session: session,
                sessionAvailable: sessionOk
            });
        }

        try {
            RT.loadConfig(function (cfg) {
                config = cfg;
                if (completed) return;
                if (cfg.failOpenTimeoutMs && cfg.failOpenTimeoutMs !== DEFAULT_TIMEOUT_MS) {
                    clearTimeout(timer);
                    timer = setTimeout(function () { finish({ allowEvent: true }, "fail_open", null); }, cfg.failOpenTimeoutMs);
                }

                // Step 1: recipients only. Most email is internal and finishes here.
                RT.getRecipients(item, function (recipients, ok) {
                    if (completed) return;
                    if (!ok) { finish({ allowEvent: true }, "fail_open", { reason: "recipients_unreadable" }); return; }
                    var quick = evaluate(recipients, [], {}, true);
                    if (quick.action === "allow") { finish({ allowEvent: true }); return; }

                    // Step 2: external recipients - read confirmations (and attachments) in parallel.
                    var attachments = [], session = {}, sessionOk = false, remaining = 2;
                    function step() {
                        if (--remaining > 0 || completed) return;
                        act(evaluate(recipients, attachments, session, sessionOk));
                    }
                    if (cfg.confirmAttachments) RT.getAttachmentNames(item, function (names) { attachments = names; step(); });
                    else step();
                    RT.getSession(item, function (values, okSession) { session = values; sessionOk = okSession; step(); });
                });
            });
        } catch (e) {
            finish({ allowEvent: true }, "fail_open", { reason: "exception" });
        }

        function act(decision) {
            var m = config.messages;
            if (decision.action === "allow") {
                finish({ allowEvent: true }, decision.reason === "confirmed" ? "sent_confirmed" : null, decision);
                return;
            }

            var externals = decision.classification.externals;
            var modern = RT.supports("1.14");

            if (decision.action === "fallback") {
                if (!modern) {
                    // Older Outlook can't offer "Send anyway" from a soft block, so fail open.
                    finish({ allowEvent: true }, "loop_breaker", decision);
                    return;
                }
                var overrideMode = (Office.MailboxEnums && Office.MailboxEnums.SendModeOverride && Office.MailboxEnums.SendModeOverride.PromptUser) || "promptUser";
                finish({
                    allowEvent: false,
                    errorMessage: Core.truncate(Core.fillTemplate(m.fallbackPrompt, {
                        count: externals.length,
                        list: Core.formatList(externals, 3)
                    }), MAX_PROMPT),
                    sendModeOverride: overrideMode
                }, "loop_breaker", decision);
                return;
            }

            // action === "block": ask the user to confirm in the task pane.
            var hint = modern
                ? Core.fillTemplate(m.hintModern, { button: m.reviewButton })
                : Core.fillTemplate(m.hintLegacy, { ribbon: m.ribbonButton });
            var pendingR = decision.pendingRecipients;
            var text = pendingR.length
                ? Core.fillTemplate(m.blockPrompt, { count: pendingR.length, list: Core.formatList(pendingR, 3), hint: hint })
                : Core.fillTemplate(m.attachmentPrompt, { attachments: decision.pendingAttachments.length, hint: hint });

            var options = { allowEvent: false, errorMessage: Core.truncate(text, MAX_PROMPT) };
            if (modern) {
                options.cancelLabel = Core.truncate(m.reviewButton, 20);
                options.commandId = REVIEW_COMMAND_ID;
            }
            RT.setSession(item, { ssBlocks: decision.blocks }, function () {
                finish(options, "prompt_shown", decision);
            });
        }
    }

    function onRecipientsChangedHandler(event) {
        var finished = false;
        function done() { if (!finished) { finished = true; clearTimeout(guard); event.completed(); } }
        var guard = setTimeout(done, 3000);

        try {
            RT.refreshConfig(false);
            var item = Office.context.mailbox.item;
            RT.loadConfig(function (cfg) {
                if (!cfg.enabled || !cfg.showRecipientBanner || !Core.isPilotUser(RT.userEmail(), cfg)) {
                    item.notificationMessages.removeAsync(BANNER_KEY, done);
                    return;
                }
                RT.getRecipients(item, function (list) {
                    var cls = Core.classifyRecipients(list, cfg);
                    if (cls.externals.length === 0) {
                        item.notificationMessages.removeAsync(BANNER_KEY, done);
                        return;
                    }
                    var message = Core.truncate(Core.fillTemplate(cfg.messages.banner, {
                        count: cls.externals.length,
                        list: Core.formatList(cls.externals, 2)
                    }), MAX_BANNER);
                    item.notificationMessages.replaceAsync(BANNER_KEY, {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: message,
                        icon: "Icon.16x16",
                        persistent: false
                    }, function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) { done(); return; }
                        // Some clients reject informational banners from event runtimes; fall back to the plain style.
                        item.notificationMessages.replaceAsync(BANNER_KEY, {
                            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                            message: message
                        }, done);
                    });
                });
            });
        } catch (e) {
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
