function showNotification(msg) {
    try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("statusAlert", {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: msg,
            icon: "Icon.16x16",
            persistent: false
        });
    } catch (e) {
        console.error("Failed to show notification: ", e);
    }
}

function onMessageSendHandler(event) {
    showNotification("ExternalSendAlert: Checking recipients...");
    // Immediately allow send for diagnostic testing
    event.completed({ allowEvent: true });
}

function onButtonClickHandler(event) {
    showNotification("ExternalSendAlert: Button was clicked!");
    event.completed();
}

// 1. Modern approach: Office.actions.associate
try {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    Office.actions.associate("onButtonClickHandler", onButtonClickHandler);
} catch (e) {
    console.error("Office.actions.associate failed", e);
}

// 2. Legacy/Fallback approach: Attach directly to global window object
// OWA sometimes requires this if the modern method fails to register in its iframe sandbox.
var g = (typeof self !== "undefined") ? self :
        (typeof window !== "undefined") ? window :
        (typeof global !== "undefined") ? global : this;

g.onMessageSendHandler = onMessageSendHandler;
g.onButtonClickHandler = onButtonClickHandler;

// Just to confirm the script evaluated successfully
console.log("ExternalSendAlert: handler_v7.js fully loaded and registered.");
