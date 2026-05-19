/*
 * External Send Alert — commands.js
 *
 * Two layers of protection:
 * 1. OnRecipientsChanged: Shows a passive yellow info banner when external recipients are added.
 * 2. OnMessageSend: Blocks send with a "Send Anyway / Don't Send" pop-up for ANY external email.
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
// HELPER: Extract external recipients from a combined list
// ============================================================================
function getExternalRecipients(allRecipients) {
    var externals = [];
    for (var i = 0; i < allRecipients.length; i++) {
        var email = allRecipients[i].emailAddress;
        if (!email || email.indexOf("@") === -1) continue;
        var domain = email.substring(email.indexOf("@") + 1).toLowerCase();
        if (INTERNAL_DOMAINS.indexOf(domain) === -1) {
            externals.push(email);
        }
    }
    return externals;
}

// ============================================================================
// HELPER: Fetch all recipients (To + Cc + Bcc) concurrently
// ============================================================================
function getAllRecipients(item) {
    var getTo  = new Promise(function(resolve) { item.to.getAsync(function(r)  { resolve(r.value || []); }); });
    var getCc  = new Promise(function(resolve) { item.cc.getAsync(function(r)  { resolve(r.value || []); }); });
    var getBcc = new Promise(function(resolve) { item.bcc.getAsync(function(r) { resolve(r.value || []); }); });
    return Promise.all([getTo, getCc, getBcc]).then(function(results) {
        return results[0].concat(results[1], results[2]);
    });
}

// ============================================================================
// HELPER: Safely fetch attachments — getAttachmentsAsync needs Mailbox 1.8.
// Returns 0 if unavailable so the handler never hangs.
// ============================================================================
function getAttachmentCount(item) {
    return new Promise(function(resolve) {
        try {
            // Check if the method exists before calling it
            if (typeof item.getAttachmentsAsync !== "function") {
                resolve(0);
                return;
            }
            // Safety timeout — resolve with 0 if it takes more than 2 seconds
            var timedOut = false;
            var timer = setTimeout(function() {
                timedOut = true;
                resolve(0);
            }, 2000);

            item.getAttachmentsAsync(function(r) {
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
    var item = Office.context.mailbox.item;

    getAllRecipients(item).then(function(allRecipients) {
        var externals = getExternalRecipients(allRecipients);

        if (externals.length > 0) {
            var externalList = externals.length <= 3
                ? externals.join(", ")
                : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";

            item.notificationMessages.replaceAsync("externalWarning", {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "External recipients detected: " + externalList,
                icon: "Icon.16x16",
                persistent: true
            });
        } else {
            item.notificationMessages.removeAsync("externalWarning");
        }

        event.completed();
    }).catch(function() {
        event.completed();
    });
}

// ============================================================================
// EVENT 2: OnMessageSend — Block pop-up for ANY email to external recipients
// ============================================================================
function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    // Run recipients and attachments in parallel, with attachments being safe/optional
    var recipientsPromise = getAllRecipients(item);
    var attachmentCountPromise = getAttachmentCount(item);

    Promise.all([recipientsPromise, attachmentCountPromise]).then(function(results) {
        var allRecipients       = results[0];
        var realAttachmentCount = results[1];
        var externals = getExternalRecipients(allRecipients);

        if (externals.length > 0) {
            var externalList = externals.length <= 3
                ? externals.join(", ")
                : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";

            var alertMessage = "EXTERNAL SEND ALERT\n\nExternal recipient(s): " + externalList + "\n\n";

            if (realAttachmentCount > 0) {
                alertMessage += "Attachments: " + realAttachmentCount + ".\n\n";
            }

            alertMessage += "Verify you are authorized to send outside the organization.";

            if (alertMessage.length > 240) {
                alertMessage = alertMessage.substring(0, 237) + "...";
            }

            event.completed({
                allowEvent: false,
                errorMessage: alertMessage
            });
        } else {
            event.completed({ allowEvent: true });
        }

    }).catch(function() {
        // Fail-open: allow send if the add-in errors out
        event.completed({ allowEvent: true });
    });
}

// ============================================================================
// RIBBON BUTTON: Passive check (no send event to block)
// ============================================================================
function onButtonClickHandler(event) {
    var item = Office.context.mailbox.item;
    getAllRecipients(item).then(function(allRecipients) {
        var externals = getExternalRecipients(allRecipients);
        if (externals.length > 0) {
            var externalList = externals.length <= 3
                ? externals.join(", ")
                : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";
            item.notificationMessages.replaceAsync("externalWarning", {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "External recipients detected: " + externalList,
                icon: "Icon.16x16",
                persistent: false
            });
        } else {
            item.notificationMessages.replaceAsync("externalClear", {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "All recipients are internal.",
                icon: "Icon.16x16",
                persistent: false
            });
        }
        event.completed();
    }).catch(function() {
        event.completed();
    });
}

// ============================================================================
// FUNCTION ASSOCIATION — Must be at top level, NOT inside Office.onReady().
// ============================================================================
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onButtonClickHandler", onButtonClickHandler);
Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);