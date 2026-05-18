/*
 * File 1: The JavaScript Engine (commands.js)
 * Purpose: Contains the OnMessageSend event handler to check for external recipients and attachments.
 */

// Initialize Office.js
Office.onReady(() => {
    // Background tasks don't typically need UI initialization
});

/**
 * Handles the OnMessageSend event.
 * Validates if the email has external recipients and attachments.
 * @param {Office.AddinCommands.Event} event
 */
function checkExternalAttachments(event) {
    let item = Office.context.mailbox.item;

    // Define your internal domains here. Add new domains to this array as needed.
    const internalDomains = ["metrixlab.com", "toluna.com"];

    // Use promises to asynchronously fetch To, Cc, and Bcc recipients
    let getTo = new Promise((resolve) => item.to.getAsync(result => resolve(result.value || [])));
    let getCc = new Promise((resolve) => item.cc.getAsync(result => resolve(result.value || [])));
    let getBcc = new Promise((resolve) => item.bcc.getAsync(result => resolve(result.value || [])));

    Promise.all([getTo, getCc, getBcc]).then(function(recipientsArrays) {
        // Combine all recipients into a single array
        let allRecipients = recipientsArrays[0].concat(recipientsArrays[1], recipientsArrays[2]);
        let isExternal = false;

        // Loop through recipients to find external domains
        for (let i = 0; i < allRecipients.length; i++) {
            let emailAddress = allRecipients[i].emailAddress;
            let domain = emailAddress.substring(emailAddress.indexOf("@") + 1).toLowerCase();
            
            // If the domain is not in our list of internal domains, it's external
            if (!internalDomains.includes(domain)) {
                isExternal = true;
                break;
            }
        }

        // Check for attachments
        item.getAttachmentsAsync(function(asyncResult) {
            let attachments = asyncResult.value || [];
            let hasAttachments = attachments.length > 0;

            if (isExternal && hasAttachments) {
                // Per requirement: Trigger notification message API
                item.notificationMessages.addAsync("ExternalAttachmentAlert", {
                    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                    message: "Security Alert: You are attempting to send an email with attachments to external recipients."
                });

                // Smart Alert block logic: stops the email from sending
                event.completed({ 
                    allowEvent: false, 
                    errorMessage: "Blocked: Sending attachments to external domains is not permitted. Please review your recipients." 
                });
            } else {
                // Allow the email to send normally
                event.completed({ allowEvent: true });
            }
        });
    }).catch(function(error) {
        // In case of error, fail safe and allow send (or block based on policy)
        console.error("Error evaluating recipients:", error);
        event.completed({ allowEvent: true });
    });
}

// Crucial step: Associate the JavaScript function with the name defined in manifest.xml
if (Office.actions && Office.actions.associate) {
    Office.actions.associate("checkExternalAttachments", checkExternalAttachments);
}
