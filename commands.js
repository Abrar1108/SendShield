/*
 * External Send Alert PoC — commands.js
 *
 * Two layers of protection:
 * 1. OnRecipientsChanged: Shows a passive yellow info banner when external recipients are added.
 * 2. OnMessageSend: Blocks send with a "Send Anyway / Don't Send" pop-up for ANY external email.
 */

// ============================================================================
// CONFIGURATION: Add or remove internal domains here as needed.
// ============================================================================
const INTERNAL_DOMAINS = [
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
        var domain = email.substring(email.indexOf("@") + 1).toLowerCase();
        if (INTERNAL_DOMAINS.indexOf(domain) === -1) {
            externals.push(email);
        }
    }
    return externals;
}

// ============================================================================
// EVENT 1: OnRecipientsChanged — Yellow info banner (non-intrusive)
// ============================================================================
function onRecipientsChangedHandler(event) {
    var item = Office.context.mailbox.item;

    var getTo  = new Promise(function(resolve) { item.to.getAsync(function(r)  { resolve(r.value || []); }); });
    var getCc  = new Promise(function(resolve) { item.cc.getAsync(function(r)  { resolve(r.value || []); }); });
    var getBcc = new Promise(function(resolve) { item.bcc.getAsync(function(r) { resolve(r.value || []); }); });

    Promise.all([getTo, getCc, getBcc]).then(function(results) {
        var allRecipients = results[0].concat(results[1], results[2]);
        var externals = getExternalRecipients(allRecipients);

        if (externals.length > 0) {
            var externalList = externals.length <= 3
                ? externals.join(", ")
                : externals.slice(0, 3).join(", ") + " (+" + (externals.length - 3) + " more)";

            // Show or update the yellow info banner
            item.notificationMessages.replaceAsync("externalWarning", {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "External recipients detected: " + externalList,
                icon: "Icon.16x16",
                persistent: true
            });
        } else {
            // Remove the banner if no external recipients remain
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
function checkExternalAttachments(event) {
    var item = Office.context.mailbox.item;

    var getTo  = new Promise(function(resolve) { item.to.getAsync(function(r)  { resolve(r.value || []); }); });
    var getCc  = new Promise(function(resolve) { item.cc.getAsync(function(r)  { resolve(r.value || []); }); });
    var getBcc = new Promise(function(resolve) { item.bcc.getAsync(function(r) { resolve(r.value || []); }); });
    var getAtt = new Promise(function(resolve) { item.getAttachmentsAsync(function(r) { resolve(r.value || []); }); });

    Promise.all([getTo, getCc, getBcc, getAtt]).then(function(results) {
        var allRecipients = results[0].concat(results[1], results[2]);
        var allAttachments = results[3];
        var externals = getExternalRecipients(allRecipients);

        if (externals.length > 0) {
            // Count real (non-inline) attachments for the message
            var realAttachments = [];
            for (var j = 0; j < allAttachments.length; j++) {
                if (!allAttachments[j].isInline) {
                    realAttachments.push(allAttachments[j]);
                }
            }

            var externalList = externals.length <= 5
                ? externals.join(", ")
                : externals.slice(0, 5).join(", ") + " (+" + (externals.length - 5) + " more)";

            var alertMessage = "EXTERNAL SEND ALERT\n\n" +
                "You are sending an email to external recipient(s):\n" +
                externalList + "\n\n";

            if (realAttachments.length > 0) {
                alertMessage += "This email contains " + realAttachments.length + " attachment(s).\n\n";
            }

            alertMessage += "Please verify you are authorized to send this email outside the organization.";

            event.completed({
                allowEvent: false,
                errorMessage: alertMessage
            });
        } else {
            // All recipients are internal — allow send
            event.completed({ allowEvent: true });
        }

    }).catch(function() {
        event.completed({ allowEvent: true });
    });
}

// ============================================================================
// FUNCTION ASSOCIATION (Required by Outlook event-based activation)
// ============================================================================
if (Office.actions && Office.actions.associate) {
    Office.actions.associate("checkExternalAttachments", checkExternalAttachments);
    Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
}
