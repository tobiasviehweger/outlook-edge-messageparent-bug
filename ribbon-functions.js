(function () {
    Office.initialize = async function (reason) {
    };
})();

function openDialog(event) {
    var dialogUrl = "https://static-resources.yasoon.com/outlook-edge-messageparent-bug/dialog.html";
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 50, displayInIframe: true }, function (asyncResult) {
        var dialog = asyncResult.value;

        //Register a message handler for the new dialog
        var i = 0;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            //Show an item message
            Office.context.mailbox.item.notificationMessages.addAsync("test" + (i++), {
                message: 'Received message, sent at ' + arg.message + ', received at ' + new Date().toLocaleString(),
                type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
            });
        });

        // Handle close event
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            event.completed();
        });
    });
}