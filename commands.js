/*
 * External Send Alert — commands.js
 *
 * Two layers of protection:
 * 1. OnRecipientsChanged: Shows a passive yellow info banner when external recipients are added.
 * 2. OnMessageSend: Blocks send with a "Send Anyway / Don't Send" pop-up for ANY external email.
 *
 * Safety: Every handler is wrapped in a master timeout (4s) to prevent
 * Outlook's "taking longer than expected" dialog from ever appearing.
 */

// ============================================================================
// CONFIGURATION: Add or remove internal domains here as needed.
// ============================================================================
var INTERNAL_DOMAINS = [
    "metrixlab.com",
    "toluna.com"
    // Add future domains below, e.g.:
    // "newentity.com",
];

// ============================================================================
// HELPER: Check if a domain is internal (supports subdomains)
// e.g. "eu.toluna.com" and "mail.metrixlab.com" are treated as internal.
// ============================================================================
function isInternalDomain(domain) {
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        if (domain === INTERNAL_DOMAINS[i] || domain.indexOf("." + INTERNAL_DOMAINS[i]) === domain.length - INTERNAL_DOMAINS[i].length - 1) {
            return true;
        }
    }
    return false;
}

// ============================================================================
// HELPER: Extract external recipients from a combined list
// ============================================================================
function getExternalRecipients(allRecipients) {
    var externals = [];
    for (var i = 0; i < allRecipients.length; i++) {
        var email = allRecipients[i].emailAddress || "";
        if (!email || email.indexOf("@") === -1) continue;
        var domain = email.substring(email.indexOf("@") + 1).toLowerCase();
        if (!isInternalDomain(domain)) {
            externals.push(email);
        }
    }
    return externals;
}

// ============================================================================
// HELPER: Fetch a single recipient field with a safety timeout
// ============================================================================
function getRecipientField(field) {
    return new Promise(function (resolve) {
        var timedOut = false;
        var timer = setTimeout(function () {
            timedOut = true;
            resolve([]);
        }, 3000);

        try {
            field.getAsync(function (r) {
                if (timedOut) return;
                clearTimeout(timer);
                resolve(r.value || []);
            });
        } catch (e) {
            if (!timedOut) {
                clearTimeout(timer);
                resolve([]);
            }
        }
    });
}

// ============================================================================
// HELPER: Fetch all recipients (To + Cc + Bcc) concurrently
// ============================================================================
function getAllRecipients(item) {
    return Promise.all([
        getRecipientField(item.to),
        getRecipientField(item.cc),
        getRecipientField(item.bcc)
    ]).then(function (results) {
        return results[0].concat(results[1], results[2]);
    });
}

// ============================================================================
// HELPER: Safely fetch attachment count — returns 0 if unavailable/timeout
// ============================================================================
function getAttachmentCount(item) {
    return new Promise(function (resolve) {
        try {
            if (typeof item.getAttachmentsAsync !== "function") {
                resolve(0);
                return;
            }
            var timedOut = false;
            var timer = setTimeout(function () {
                timedOut = true;
                resolve(0);
            }, 2000);

            item.getAttachmentsAsync(function (r) {
                if (timedOut) return;
                clearTimeout(timer);
                var attachments = r.value || [];
                var count = 0;
                for (var j = 0; j < attachments.length; j++) {
                    if (!attachments[j].isInline) count++;
                }
                resolve(count);
            });
        } catch (e) {
            resolve(0);
        }
    });
}

// ============================================================================
// EVENT 1: OnRecipientsChanged — Yellow info banner (non-intrusive)
// ============================================================================
function onRecipientsChangedHandler(event) {
    // Master timeout: always complete within 4 seconds
    var completed = false;
    var finish = function () {
        if (!completed) {
            completed = true;
            event.completed();
        }
    };
    setTimeout(finish, 4000);

    try {
        var item = Office.context.mailbox.item;

        getAllRecipients(item).then(function (allRecipients) {
            var externals = getExternalRecipients(allRecipients);

            if (externals.length > 0) {
                var externalList = externals.length <= 3
                    ? externals.join(", ")
                    : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";

                item.notificationMessages.replaceAsync("externalWarning", {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: "\u26A0 External recipients detected: " + externalList,
                    icon: "Icon.16x16",
                    persistent: true
                });
            } else {
                item.notificationMessages.removeAsync("externalWarning");
            }

            finish();
        }).catch(function () {
            finish();
        });
    } catch (e) {
        finish();
    }
}

// ============================================================================
// EVENT 2: OnMessageSend — Block pop-up for external recipients only
// Internal-only emails pass through instantly with zero delay.
// ============================================================================
function onMessageSendHandler(event) {
    // Master timeout: always complete within 4 seconds (fail-open)
    var completed = false;
    var finish = function (options) {
        if (!completed) {
            completed = true;
            event.completed(options || { allowEvent: true });
        }
    };
    setTimeout(function () { finish({ allowEvent: true }); }, 4000);

    try {
        var item = Office.context.mailbox.item;

        // Run recipients and attachments in parallel
        var recipientsPromise = getAllRecipients(item);
        var attachmentCountPromise = getAttachmentCount(item);

        Promise.all([recipientsPromise, attachmentCountPromise]).then(function (results) {
            var allRecipients = results[0];
            var realAttachmentCount = results[1];
            var externals = getExternalRecipients(allRecipients);

            // ── Internal only: allow send immediately, no popup ──
            if (externals.length === 0) {
                finish({ allowEvent: true });
                return;
            }

            // ── External recipients found: build alert and block ──
            var externalList = externals.length <= 3
                ? externals.join(", ")
                : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";

            var alertMessage = "EXTERNAL SEND ALERT\n\n"
                + "External recipient(s): " + externalList + "\n\n";

            if (realAttachmentCount > 0) {
                alertMessage += "Attachments: " + realAttachmentCount + " file(s) included.\n\n";
            }

            alertMessage += "Verify you are authorized to send outside the organization.";

            // Safety trim to avoid Outlook truncation (max ~250 chars)
            if (alertMessage.length > 240) {
                alertMessage = alertMessage.substring(0, 237) + "...";
            }

            finish({
                allowEvent: false,
                errorMessage: alertMessage
            });

        }).catch(function () {
            // Fail-open: allow send if the add-in errors out
            finish({ allowEvent: true });
        });
    } catch (e) {
        // Fail-open on any unexpected error
        finish({ allowEvent: true });
    }
}

// ============================================================================
// RIBBON BUTTON: Manual "Check External" button
// ============================================================================
function onButtonClickHandler(event) {
    var completed = false;
    var finish = function () {
        if (!completed) {
            completed = true;
            event.completed();
        }
    };
    setTimeout(finish, 4000);

    try {
        var item = Office.context.mailbox.item;
        getAllRecipients(item).then(function (allRecipients) {
            var externals = getExternalRecipients(allRecipients);
            if (externals.length > 0) {
                var externalList = externals.length <= 3
                    ? externals.join(", ")
                    : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";
                item.notificationMessages.replaceAsync("externalWarning", {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: "\u26A0 External recipients: " + externalList,
                    icon: "Icon.16x16",
                    persistent: false
                });
            } else {
                item.notificationMessages.replaceAsync("externalClear", {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: "\u2705 All recipients are internal.",
                    icon: "Icon.16x16",
                    persistent: false
                });
            }
            finish();
        }).catch(function () {
            finish();
        });
    } catch (e) {
        finish();
    }
}

// ============================================================================
// FUNCTION ASSOCIATION — Must be at top level, NOT inside Office.onReady().
// ============================================================================
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onButtonClickHandler", onButtonClickHandler);
Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);