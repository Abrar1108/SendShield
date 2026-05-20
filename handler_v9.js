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

function formatList(arr) {
    if (arr.length <= 3) return arr.join(", ");
    return arr.slice(0, 3).join(", ") + " (+" + (arr.length - 3) + " more)";
}

function blockWithMessage(event, externals, attachmentCount) {
    var msg = "External recipients: " + formatList(externals) + ".\n";
    if (attachmentCount > 0) {
        msg += "Attachments included: " + attachmentCount + " file(s).\n";
    }
    msg += "Are you sure you want to send outside the organization?";

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
    var item = Office.context.mailbox.item;
    
    // Safety timeout: Never block Outlook for more than 4.5 seconds
    var isCompleted = false;
    var safetyTimeout = setTimeout(function() {
        if (!isCompleted) {
            isCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, 4500);
    
    item.to.getAsync({ asyncContext: { event: event, all: [] } }, function (toResult) {
        if (isCompleted) return;
        var ctx = toResult.asyncContext;
        ctx.all = ctx.all.concat(toResult.value || []);
        
        item.cc.getAsync({ asyncContext: ctx }, function (ccResult) {
            if (isCompleted) return;
            var ctx2 = ccResult.asyncContext;
            ctx2.all = ctx2.all.concat(ccResult.value || []);
            
            item.bcc.getAsync({ asyncContext: ctx2 }, function (bccResult) {
                if (isCompleted) return;
                var ctx3 = bccResult.asyncContext;
                ctx3.all = ctx3.all.concat(bccResult.value || []);
                
                var externals = getExternalsFromList(ctx3.all);
                
                if (externals.length === 0) {
                    isCompleted = true;
                    clearTimeout(safetyTimeout);
                    ctx3.event.completed({ allowEvent: true });
                } else {
                    if (typeof item.getAttachmentsAsync === "function") {
                        item.getAttachmentsAsync(function(attResult) {
                            if (isCompleted) return;
                            isCompleted = true;
                            clearTimeout(safetyTimeout);
                            
                            var attachments = attResult.value || [];
                            var realCount = 0;
                            for (var j = 0; j < attachments.length; j++) {
                                if (!attachments[j].isInline) realCount++;
                            }
                            blockWithMessage(ctx3.event, externals, realCount);
                        });
                    } else {
                        if (isCompleted) return;
                        isCompleted = true;
                        clearTimeout(safetyTimeout);
                        blockWithMessage(ctx3.event, externals, 0);
                    }
                }
            });
        });
    });
}

// -------------------------------------------------------------
// EVENT HANDLER: On Recipients Changed (Shows non-intrusive banner)
// -------------------------------------------------------------
function onRecipientsChangedHandler(event) {
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
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "Notice: You are sending to external recipients (" + formatList(externals) + ").",
                        icon: "Icon.16x16",
                        persistent: false
                    });
                } else {
                    item.notificationMessages.removeAsync("extWarning");
                }
                event.completed();
            });
        });
    });
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
