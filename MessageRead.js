
Office.initialize = function (reason) {
  $(document).ready(function () {
    var outlookItem = Office.context.mailbox.item;

    // Build query parameters from the selected email
    var parameters =
      "&messageId=" + encodeURIComponent(outlookItem.itemId) +
      "&subject=" + encodeURIComponent(outlookItem.subject) +
      "&from=" + encodeURIComponent(outlookItem.from.emailAddress) +
      "&fromname=" + encodeURIComponent(outlookItem.from.displayName) +
      "&receivedOn=" + encodeURIComponent(outlookItem.dateTimeCreated);

    // Append parameters to the iframe URL
    document.getElementById('myApp').src += parameters;
  });
};
