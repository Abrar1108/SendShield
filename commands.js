/*
 * External Send Alert — DIAGNOSTIC TEST (v6.1)
 * Minimal handler — does nothing except immediately allow send.
 * If this still shows "taking longer than expected", the JS file is not loading.
 */

function onMessageSendHandler(event) {
    event.completed({ allowEvent: true });
}

function onButtonClickHandler(event) {
    event.completed();
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onButtonClickHandler", onButtonClickHandler);