// Office.onReady is MANDATORY for event-based activation.
Office.onReady(function (info) {});

var INTERNAL_DOMAINS = [
    "metrixlab.com",
    "toluna.com"
];

function isInternalDomain(domain) {
    domain = domain.toLowerCase();
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        var d = INTERNAL_DOMAINS[i];
        if (domain === d || (domain.length > d.length && domain.indexOf("." + d) === domain.length - d.length - 1)) {
            return true;
        }
    }
    return false;
}

function getExternalsFromList(recipients) {
    var externals = [];
    for (var i = 0; i < recipients.length; i++) {
        var email = recipients[i].emailAddress || "";
        if (!email || email.indexOf("@") === -1) continue;
        var domain = email.substring(email.indexOf("@") + 1);
        if (!isInternalDomain(domain)) {
            externals.push(email);
        }
    }
    return externals;
}

// Limits the display to 3 items, then adds "(+X more)"
function formatList(arr) {
    if (arr.length <= 3) return arr.join(", ");
    return arr.slice(0, 3).join(", ") + " (+" + (arr.length - 3) + " more)";
}

function blockWithMessage(event, externals, attachmentNames) {
    var msg = "External recipients: " + formatList(externals) + ".\n";
    if (attachmentNames && attachmentNames.length > 0) {
        msg += "Attachments (" + attachmentNames.length + "): " + formatList(attachmentNames) + ".\n";
    }
    msg += "Are you sure you want to send outside the organization?";

    // Outlook strictly limits the prompt message to 250 characters.
    if (msg.length > 240) {
        msg = msg.substring(0, 237) + "...";
    }

    event.completed({
        allowEvent: false,
        errorMessage: msg
    });
}

// -------------------------------------------------------------
// EVENT HANDLER: On Message Send (Blocks/Prompts if External)
// -------------------------------------------------------------
function onMessageSendHandler(event) {
    var isCompleted = false;
    var currentExternals = [];
    
    // Safety timeout: Never block Outlook for more than 3.5 seconds
    var safetyTimeout = setTimeout(function() {
        if (!isCompleted) {
            isCompleted = true;
            // If we already detected externals but attachments hung, just block them.
            if (currentExternals.length > 0) {
                blockWithMessage(event, currentExternals, []);
            } else {
                event.completed({ allowEvent: true });
            }
        }
    }, 3500);
    
    try {
        var item = Office.context.mailbox.item;
        item.to.getAsync({ asyncContext: {} }, function (toResult) {
            if (isCompleted) return;
            var allRecips = (toResult.value || []);
            
            item.cc.getAsync({ asyncContext: {} }, function (ccResult) {
                if (isCompleted) return;
                allRecips = allRecips.concat(ccResult.value || []);
                
                item.bcc.getAsync({ asyncContext: {} }, function (bccResult) {
                    if (isCompleted) return;
                    allRecips = allRecips.concat(bccResult.value || []);
                    
                    currentExternals = getExternalsFromList(allRecips);
                    
                    if (currentExternals.length === 0) {
                        isCompleted = true;
                        clearTimeout(safetyTimeout);
                        event.completed({ allowEvent: true });
                    } else {
                        if (typeof item.getAttachmentsAsync === "function") {
                            try {
                                item.getAttachmentsAsync(function(attResult) {
                                    if (isCompleted) return;
                                    isCompleted = true;
                                    clearTimeout(safetyTimeout);
                                    
                                    var attachmentNames = [];
                                    if (attResult.status === Office.AsyncResultStatus.Succeeded) {
                                        var attachments = attResult.value || [];
                                        for (var j = 0; j < attachments.length; j++) {
                                            if (!attachments[j].isInline) {
                                                attachmentNames.push(attachments[j].name);
                                            }
                                        }
                                    }
                                    blockWithMessage(event, currentExternals, attachmentNames);
                                });
                            } catch(e) {
                                if (isCompleted) return;
                                isCompleted = true;
                                clearTimeout(safetyTimeout);
                                blockWithMessage(event, currentExternals, []);
                            }
                        } else {
                            if (isCompleted) return;
                            isCompleted = true;
                            clearTimeout(safetyTimeout);
                            blockWithMessage(event, currentExternals, []);
                        }
                    }
                });
            });
        });
    } catch(e) {
        if (!isCompleted) {
            isCompleted = true;
            clearTimeout(safetyTimeout);
            event.completed({ allowEvent: true });
        }
    }
}

// -------------------------------------------------------------
// EVENT HANDLER: On Recipients Changed (Shows non-intrusive banner)
// -------------------------------------------------------------
function onRecipientsChangedHandler(event) {
    try {
        var item = Office.context.mailbox.item;
        item.to.getAsync(function (toResult) {
            var toList = toResult.value || [];
            item.cc.getAsync(function (ccResult) {
                var ccList = ccResult.value || [];
                item.bcc.getAsync(function (bccResult) {
                    var bccList = bccResult.value || [];
                    var all = toList.concat(ccList, bccList);
                    var externals = getExternalsFromList(all);
                    
                    if (externals.length > 0) {
                        item.notificationMessages.replaceAsync("extWarning", {
                            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                            message: "Notice: You are sending to external recipients (" + formatList(externals) + ").",
                            icon: "Shield.16x16",
                            persistent: false
                        }, function(result) {
                            event.completed();
                        });
                    } else {
                        item.notificationMessages.removeAsync("extWarning", function(result) {
                            event.completed();
                        });
                    }
                });
            });
        });
    } catch (e) {
        event.completed();
    }
}

function onButtonClickHandler(event) {
    event.completed();
}

// Register Handlers (Modern + Fallback)
try {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
    Office.actions.associate("onButtonClickHandler", onButtonClickHandler);
} catch (e) {}

var g = (typeof self !== "undefined") ? self :
        (typeof window !== "undefined") ? window :
        (typeof global !== "undefined") ? global : this;
g.onMessageSendHandler = onMessageSendHandler;
g.onRecipientsChangedHandler = onRecipientsChangedHandler;
g.onButtonClickHandler = onButtonClickHandler;
